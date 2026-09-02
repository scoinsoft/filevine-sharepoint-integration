const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { sharepoint } = require('../config/env');
const { log, logError } = require('../utils/logger');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const DEFAULT_TIMEOUT_MS = 180000;
const SIMPLE_UPLOAD_MAX_BYTES = 4 * 1024 * 1024;
// Graph requires multiples of 320 KiB; max fragment is 60 MiB.
// 40 MiB balances throughput vs memory under SYNC_CONCURRENCY parallel uploads.
const UPLOAD_SESSION_CHUNK_BYTES = 40 * 1024 * 1024;
const MAX_RETRIES = 5;

let cachedToken = null;
let tokenExpiresAt = 0;

function sanitizePathSegment(name) {
  return String(name || '')
    .replace(/[~"#%&*:<>?/\\{|}]/g, '_')
    .trim() || 'Unnamed';
}

function sanitizeFolderName(name) {
  return sanitizePathSegment(name);
}

function getRootFolder() {
  const value = sharepoint.rootFolder();
  return sanitizePathSegment(value || 'Filevine');
}

function buildSharePointPath(projectName, filename) {
  const safeProjectName = sanitizeFolderName(projectName);
  const safeFilename = sanitizePathSegment(path.basename(filename));
  return `${getRootFolder()}/${safeProjectName}/${safeFilename}`;
}

function buildProjectFolderPath(projectName) {
  const safeProjectName = sanitizeFolderName(projectName);
  return `${getRootFolder()}/${safeProjectName}`;
}

function isItemNotFoundError(error) {
  if (!error) return false;
  const status = error?.response?.status;
  const code = getGraphErrorCode(error);
  if (status === 404) return true;
  return code === 'itemNotFound';
}

/** Serialize uploads targeting the same SharePoint path. */
const uploadPathLocks = new Map();

async function withUploadPathLock(sharePointPath, operation) {
  const previous = uploadPathLocks.get(sharePointPath) || Promise.resolve();
  let release = () => {};
  const current = previous.then(
    () =>
      new Promise((resolve) => {
        release = resolve;
      })
  );
  uploadPathLocks.set(sharePointPath, current);

  await previous;
  try {
    return await operation();
  } finally {
    release();
    if (uploadPathLocks.get(sharePointPath) === current) {
      uploadPathLocks.delete(sharePointPath);
    }
  }
}

function getUploadTimeoutMs() {
  const { isServerless, remainingMs } = require('../config/runtime');
  const ms = Number(
    process.env.SHAREPOINT_TIMEOUT_MS ||
      process.env.POWER_AUTOMATE_TIMEOUT_MS ||
      DEFAULT_TIMEOUT_MS
  );
  const configured = Number.isFinite(ms) && ms > 0 ? Math.floor(ms) : DEFAULT_TIMEOUT_MS;
  if (!isServerless()) {
    return configured;
  }
  const remaining = remainingMs();
  if (!Number.isFinite(remaining)) {
    return Math.min(configured, 120000);
  }
  return Math.max(5000, Math.min(configured, remaining - 5000));
}

function formatMegabytes(bytes) {
  return (Number(bytes) / (1024 * 1024)).toFixed(1);
}

function buildUploadProgressPayload(filename, bytesUploaded, bytesTotal) {
  const total = Math.max(0, Number(bytesTotal) || 0);
  const uploaded = Math.max(0, Number(bytesUploaded) || 0);
  const percent = total > 0 ? Math.min(100, Math.round((uploaded / total) * 100)) : 0;
  return {
    filename,
    bytesUploaded: uploaded,
    bytesTotal: total,
    percent,
    stage: 'uploading',
  };
}

function reportUploadProgress(onProgress, filename, bytesUploaded, bytesTotal) {
  if (typeof onProgress !== 'function') return;
  onProgress(buildUploadProgressPayload(filename, bytesUploaded, bytesTotal));
}

function isResourceModifiedError(error) {
  if (!error) return false;
  if (error.resourceModified) return true;
  const code = getGraphErrorCode(error);
  if (code === 'resourceModified') return true;
  const message = String(error.message || '');
  return (
    message.includes('resourceModified') ||
    message.includes('eTag mismatch') ||
    (message.includes('status 409') && message.includes('resourceModified'))
  );
}

function isSharePointConflictSkipError(error) {
  return isNameAlreadyExistsError(error) || isResourceModifiedError(error);
}

function isRetryableUploadError(error) {
  if (isSharePointConfigError(error)) return false;
  if (isNameAlreadyExistsError(error)) return false;
  if (isResourceModifiedError(error)) return true;

  const status = error?.response?.status;
  if ([408, 429, 500, 502, 503, 504].includes(status)) return true;

  const code = error?.code;
  if (['ECONNRESET', 'ETIMEDOUT', 'ECONNABORTED', 'ENOTFOUND', 'EAI_AGAIN', 'ENOENT'].includes(code)) {
    return true;
  }

  const message = String(error?.message || '');
  return (
    message.includes('Failed to read upload chunk') ||
    message.includes('Local file not found') ||
    message.includes('ENOENT') ||
    message.includes('socket hang up')
  );
}

function isUnauthorizedError(error) {
  return error?.response?.status === 401;
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function clearCachedToken() {
  cachedToken = null;
  tokenExpiresAt = 0;
}

function encodeGraphPath(relativePath) {
  return String(relativePath || '')
    .split('/')
    .filter(Boolean)
    .map((segment) => encodeURIComponent(segment))
    .join('/');
}

function buildItemContentUrl(itemPath) {
  const siteId = encodeURIComponent(sharepoint.siteId());
  const driveId = encodeURIComponent(sharepoint.driveId());
  const encodedPath = encodeGraphPath(itemPath);
  return `${GRAPH_BASE}/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/content`;
}

function buildItemByPathUrl(itemPath) {
  const siteId = encodeURIComponent(sharepoint.siteId());
  const driveId = encodeURIComponent(sharepoint.driveId());
  const encodedPath = encodeGraphPath(itemPath);
  return `${GRAPH_BASE}/sites/${siteId}/drives/${driveId}/root:/${encodedPath}`;
}

function buildItemUploadSessionUrl(itemPath) {
  const siteId = encodeURIComponent(sharepoint.siteId());
  const driveId = encodeURIComponent(sharepoint.driveId());
  const encodedPath = encodeGraphPath(itemPath);
  return `${GRAPH_BASE}/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/createUploadSession`;
}

async function getAccessToken() {
  if (cachedToken && Date.now() < tokenExpiresAt - 60_000) {
    return cachedToken;
  }

  const tenantId = sharepoint.tenantId();
  const clientId = sharepoint.clientId();
  const clientSecret = sharepoint.clientSecret();

  try {
    const response = await axios.post(
      `https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`,
      new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials',
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        timeout: getUploadTimeoutMs(),
      }
    );

    if (!response.data?.access_token) {
      throw createSharePointConfigError(
        'SharePoint authentication failed: Azure did not return an access token.'
      );
    }

    const expiresIn = Number(response.data.expires_in) || 3600;
    cachedToken = response.data.access_token;
    tokenExpiresAt = Date.now() + expiresIn * 1000;
    return cachedToken;
  } catch (error) {
    const configError = parseAzureAuthError(error);
    if (configError) {
      throw configError;
    }
    throw formatGraphError(error, 'authentication');
  }
}

async function verifySharePointAuth() {
  clearCachedToken();
  await getAccessToken();
  log('SharePoint authentication verified');
}

function createSharePointConfigError(message, cause) {
  const error = new Error(message);
  error.name = 'SharePointConfigError';
  error.sharePointConfigError = true;
  if (cause) {
    error.cause = cause;
  }
  return error;
}

function parseAzureAuthError(error) {
  const data = error?.response?.data;
  const raw =
    typeof data === 'string'
      ? data
      : data?.error_description || data?.error || JSON.stringify(data || {});
  const text = String(raw || error?.message || '');

  if (text.includes('AADSTS90002') || text.includes('AADSTS900023')) {
    return createSharePointConfigError(
      'SharePoint is misconfigured: Azure tenant ID/domain is invalid. Use your exact tenant GUID from Azure Portal or your *.onmicrosoft.com domain in AZURE_TENANT_ID, then restart the server.',
      error
    );
  }
  if (text.includes('AADSTS7000215') || text.includes('invalid_client')) {
    return createSharePointConfigError(
      'SharePoint is misconfigured: Azure client secret is invalid. Check AZURE_CLIENT_SECRET in .env.',
      error
    );
  }
  if (text.includes('AADSTS700016') || text.includes('was not found in the directory')) {
    return createSharePointConfigError(
      'SharePoint is misconfigured: Azure client ID was not found in the tenant. Check AZURE_CLIENT_ID in .env.',
      error
    );
  }
  if (error?.response?.status === 400 || error?.response?.status === 401) {
    return createSharePointConfigError(
      `SharePoint authentication failed. Verify AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in .env, then restart the server. Details: ${text}`,
      error
    );
  }
  return null;
}

function isSharePointConfigError(error) {
  if (!error) return false;
  if (error.sharePointConfigError) return true;
  const message = String(error.message || '');
  return (
    message.includes('SharePoint is misconfigured') ||
    message.includes('SharePoint authentication failed') ||
    message.includes('AADSTS')
  );
}

function getGraphErrorCode(error) {
  const data = error?.response?.data;
  if (data && typeof data === 'object') {
    return data.error?.code || data.code || null;
  }
  const message = String(error?.message || '');
  const match = message.match(/"code"\s*:\s*"([^"]+)"/);
  return match ? match[1] : null;
}

function isNameAlreadyExistsError(error) {
  if (!error) return false;
  if (error.nameAlreadyExists) return true;

  const status = error?.response?.status;
  const code = getGraphErrorCode(error);
  if (status === 409 && code === 'nameAlreadyExists') return true;

  const message = String(error.message || '');
  return (
    message.includes('nameAlreadyExists') ||
    message.includes('A file with the same name is currently being uploaded') ||
    (message.includes('status 409') && message.includes('nameAlreadyExists'))
  );
}

function formatGraphError(error, action) {
  const configError = parseAzureAuthError(error);
  if (configError) {
    return configError;
  }

  if (error.sharePointConfigError) {
    return error;
  }

  if (isNameAlreadyExistsError(error)) {
    const conflictError = new Error(
      `SharePoint ${action} skipped: a file with the same name already exists or is currently uploading`
    );
    conflictError.nameAlreadyExists = true;
    conflictError.status = error?.response?.status || 409;
    return conflictError;
  }

  if (isResourceModifiedError(error)) {
    const conflictError = new Error(
      `SharePoint ${action} skipped: file already exists or was modified by another upload`
    );
    conflictError.resourceModified = true;
    conflictError.status = error?.response?.status || 409;
    return conflictError;
  }

  if (error.response) {
    const detail =
      typeof error.response.data === 'string'
        ? error.response.data
        : JSON.stringify(error.response.data);
    return new Error(`SharePoint ${action} failed (status ${error.response.status}): ${detail}`);
  }
  return error;
}

async function withAccessToken(operation) {
  try {
    const token = await getAccessToken();
    return await operation(token);
  } catch (error) {
    if (isSharePointConfigError(error)) {
      throw error;
    }
    if (!isUnauthorizedError(error)) {
      throw error;
    }
    clearCachedToken();
    const token = await getAccessToken();
    return operation(token);
  }
}

async function uploadSmallFile(accessToken, itemPath, buffer, contentType, onProgress, filename) {
  const url = buildItemContentUrl(itemPath);
  reportUploadProgress(onProgress, filename, 0, buffer.length);
  const response = await axios.put(url, buffer, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': contentType || 'application/octet-stream',
    },
    timeout: getUploadTimeoutMs(),
    maxBodyLength: Infinity,
    maxContentLength: Infinity,
    validateStatus: (status) => status >= 200 && status < 300,
    onUploadProgress: (event) => {
      reportUploadProgress(onProgress, filename, event.loaded, event.total || buffer.length);
    },
  });
  reportUploadProgress(onProgress, filename, buffer.length, buffer.length);
  return response.data;
}

async function createUploadSession(accessToken, itemPath) {
  const sessionUrl = buildItemUploadSessionUrl(itemPath);
  const sessionResponse = await axios.post(
    sessionUrl,
    {
      item: {
        '@microsoft.graph.conflictBehavior': 'replace',
        name: path.basename(itemPath),
      },
    },
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      timeout: getUploadTimeoutMs(),
    }
  );

  const uploadUrl = sessionResponse.data?.uploadUrl;
  if (!uploadUrl) {
    throw new Error('SharePoint upload session failed: missing uploadUrl');
  }
  return uploadUrl;
}

async function uploadLargeFileFromPath(accessToken, itemPath, filePath, fileSize, contentType, onProgress, filename) {
  if (!fs.existsSync(filePath)) {
    const missing = new Error(`Local file not found for upload: ${filePath}`);
    missing.code = 'ENOENT';
    throw missing;
  }

  const liveStats = fs.statSync(filePath);
  const liveSize = liveStats.size;
  if (liveSize !== fileSize) {
    log('Local file size changed before upload; using current size', {
      filename,
      expectedSize: fileSize,
      actualSize: liveSize,
      filePath,
    });
  }
  const totalSize = liveSize;

  let sessionAttempts = 0;
  const maxSessionAttempts = 3;

  while (sessionAttempts < maxSessionAttempts) {
    sessionAttempts += 1;
    const uploadUrl = await createUploadSession(accessToken, itemPath);
    const handle = await fs.promises.open(filePath, 'r');
    let offset = 0;
    let result = null;
    let lastReportedAt = 0;
    let sessionError = null;

    try {
      reportUploadProgress(onProgress, filename, 0, totalSize);

      while (offset < totalSize) {
        const chunkEnd = Math.min(offset + UPLOAD_SESSION_CHUNK_BYTES, totalSize);
        const chunkSize = chunkEnd - offset;
        const chunk = Buffer.allocUnsafe(chunkSize);
        const { bytesRead } = await handle.read(chunk, 0, chunkSize, offset);
        if (bytesRead !== chunkSize) {
          const err = new Error(
            `Failed to read upload chunk at offset ${offset} (expected ${chunkSize}, got ${bytesRead})`
          );
          err.code = 'CHUNK_READ_MISMATCH';
          throw err;
        }

        const contentRange = `bytes ${offset}-${chunkEnd - 1}/${totalSize}`;
        try {
          const response = await axios.put(uploadUrl, chunk.subarray(0, chunkSize), {
            headers: {
              'Content-Length': chunkSize,
              'Content-Range': contentRange,
              'Content-Type': contentType || 'application/octet-stream',
            },
            timeout: getUploadTimeoutMs(),
            maxBodyLength: Infinity,
            maxContentLength: Infinity,
            validateStatus: (status) => status === 200 || status === 201 || status === 202,
          });

          if (response.status === 200 || response.status === 201) {
            result = response.data;
            reportUploadProgress(onProgress, filename, totalSize, totalSize);
            break;
          }

          offset = chunkEnd;
          const now = Date.now();
          if (now - lastReportedAt >= 400 || offset >= totalSize) {
            lastReportedAt = now;
            reportUploadProgress(onProgress, filename, offset, totalSize);
          }
        } catch (chunkError) {
          if (isResourceModifiedError(chunkError) && sessionAttempts < maxSessionAttempts) {
            sessionError = chunkError;
            log('SharePoint chunk conflict; restarting upload session', {
              filename,
              itemPath,
              offset,
              attempt: sessionAttempts,
            });
            break;
          }
          throw chunkError;
        }
      }
    } finally {
      await handle.close().catch(() => {});
    }

    if (result) {
      return result;
    }
    if (sessionError) {
      await delay(1000 * sessionAttempts);
      continue;
    }
    break;
  }

  throw new Error('SharePoint upload session did not complete');
}

async function uploadFileToGraph(itemPath, filePath, fileSize, contentType, onProgress, filename) {
  return withAccessToken(async (accessToken) => {
    if (fileSize <= SIMPLE_UPLOAD_MAX_BYTES) {
      const buffer = await fs.promises.readFile(filePath);
      return uploadSmallFile(accessToken, itemPath, buffer, contentType, onProgress, filename);
    }
    // Stream from disk in chunks — avoids Node's ~2 GiB Buffer limit.
    return uploadLargeFileFromPath(
      accessToken,
      itemPath,
      filePath,
      fileSize,
      contentType,
      onProgress,
      filename
    );
  });
}

async function uploadToSharePoint({
  projectId,
  projectName,
  filename,
  contentType,
  filePath,
  onProgress,
}) {
  const sharePointPath = buildSharePointPath(projectName, filename);
  if (!filePath) {
    throw new Error(`Missing local file path for upload: ${filename}`);
  }
  if (!fs.existsSync(filePath)) {
    throw new Error(`Local file not found for upload: ${filePath}`);
  }

  const stats = fs.statSync(filePath);
  const fileSize = stats.size;

  log('Uploading to SharePoint via Microsoft Graph', {
    projectId,
    filename,
    sharePointPath,
    contentType,
    size: fileSize,
    timeoutMs: getUploadTimeoutMs(),
  });

  try {
    let lastError;

    for (let attempt = 1; attempt <= MAX_RETRIES; attempt += 1) {
      try {
        const graphItem = await withUploadPathLock(sharePointPath, () =>
          uploadFileToGraph(
            sharePointPath,
            filePath,
            fileSize,
            contentType || 'application/octet-stream',
            onProgress,
            filename
          )
        );

        log('Upload successful', {
          projectId,
          filename,
          sharePointPath,
          graphItemId: graphItem?.id || null,
        });

        return {
          success: true,
          message: 'File uploaded to SharePoint',
          sharePointPath,
          filename,
          projectId,
          projectName,
          graphItemId: graphItem?.id || null,
          webUrl: graphItem?.webUrl || null,
        };
      } catch (attemptError) {
        lastError = attemptError;
        if (isResourceModifiedError(attemptError) && attempt >= 2) {
          throw attemptError;
        }
        if (attempt < MAX_RETRIES && isRetryableUploadError(attemptError)) {
          const backoffMs = 1000 * attempt;
          log('Retrying SharePoint upload after error', {
            projectId,
            filename,
            sharePointPath,
            attempt,
            maxAttempts: MAX_RETRIES,
            code: attemptError.code || getGraphErrorCode(attemptError),
            status: attemptError?.response?.status || null,
            message: attemptError.message,
            backoffMs,
            resourceModified: isResourceModifiedError(attemptError),
          });
          await delay(backoffMs);
          continue;
        }
        throw attemptError;
      }
    }

    throw lastError || new Error('SharePoint upload failed: no response received');
  } catch (error) {
    if (isSharePointConflictSkipError(error)) {
      log('SharePoint upload conflict; treating as already uploaded', {
        projectId,
        filename,
        sharePointPath,
        code: getGraphErrorCode(error) || 'conflict',
      });
    } else {
      logError('SharePoint upload failed', error);
      log('Upload debug context', { projectId, filename, sharePointPath });
    }
    throw formatGraphError(error, 'upload');
  }
}

async function uploadFile(filePath, projectId, projectName, contentType, options = {}) {
  const filename = options.filename || path.basename(filePath);

  return uploadToSharePoint({
    projectId,
    projectName,
    filename,
    contentType: contentType || 'application/octet-stream',
    filePath,
    onProgress: typeof options.onProgress === 'function' ? options.onProgress : null,
  });
}

async function deleteFolderByPath(relativeFolderPath) {
  const folderPath = String(relativeFolderPath || '').replace(/^\/+/, '');
  if (!folderPath) {
    throw new Error('SharePoint delete failed: missing folder path');
  }

  return withAccessToken(async (accessToken) => {
    const url = `${buildItemByPathUrl(folderPath)}:`;
    try {
      await axios.delete(url, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        timeout: getUploadTimeoutMs(),
        validateStatus: (status) => status === 204 || status === 200,
      });
      return { deleted: true, existed: true, folderPath };
    } catch (error) {
      if (isItemNotFoundError(error)) {
        return { deleted: false, existed: false, notFound: true, folderPath };
      }
      throw formatGraphError(error, 'delete');
    }
  });
}

const STATE_FOLDER_NAME = '_sync-state';

function getStateRootFolder() {
  return `${getRootFolder()}/${STATE_FOLDER_NAME}`;
}

function buildStateItemPath(relativePath) {
  const normalized = String(relativePath || '')
    .replace(/\\/g, '/')
    .replace(/^\/+/, '');
  return `${getStateRootFolder()}/${normalized}`;
}

async function graphGet(accessToken, url, options = {}) {
  return axios.get(url, {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
    },
    timeout: getUploadTimeoutMs(),
    validateStatus: options.validateStatus,
    responseType: options.responseType || 'json',
    transformResponse: options.transformResponse,
  });
}

async function listFolderChildren(relativeFolderPath) {
  const folderPath = String(relativeFolderPath || '').replace(/^\/+/, '');
  if (!folderPath) {
    return [];
  }

  return withAccessToken(async (accessToken) => {
    const children = [];
    let url = `${buildItemByPathUrl(folderPath)}:/children?$select=name,folder,file,lastModifiedDateTime,size&$top=200`;

    try {
      while (url) {
        const response = await graphGet(accessToken, url, {
          validateStatus: (status) => status >= 200 && status < 300,
        });
        const items = Array.isArray(response.data?.value) ? response.data.value : [];
        children.push(...items);
        url = response.data?.['@odata.nextLink'] || null;
      }
      return children;
    } catch (error) {
      if (isItemNotFoundError(error)) {
        return [];
      }
      throw formatGraphError(error, 'list');
    }
  });
}

async function listFolderFileNames(relativeFolderPath) {
  const children = await listFolderChildren(relativeFolderPath);
  return children.filter((item) => item?.file).map((item) => String(item.name || '')).filter(Boolean);
}

const FOLDER_NAME_CACHE_RELATIVE = 'sharepoint-project-folders.json';
const ROOT_FOLDER_CACHE_MS = 6 * 60 * 60 * 1000;

let cachedRootFolderNames = null;
let cachedRootFolderNamesAt = 0;

function folderNameSetFromList(names) {
  const set = new Set();
  for (const name of names || []) {
    const value = String(name || '').trim().toLowerCase();
    if (value && !value.startsWith('_') && !value.startsWith('.')) {
      set.add(value);
    }
  }
  return set;
}

async function getRootProjectFolderNameSet({ forceRefresh = false } = {}) {
  if (
    !forceRefresh &&
    cachedRootFolderNames &&
    Date.now() - cachedRootFolderNamesAt < ROOT_FOLDER_CACHE_MS
  ) {
    return cachedRootFolderNames;
  }

  if (!forceRefresh) {
    try {
      const stored = await readStateJson(FOLDER_NAME_CACHE_RELATIVE);
      const storedAt = stored?.updatedAt ? new Date(stored.updatedAt).getTime() : 0;
      if (Array.isArray(stored?.names) && Date.now() - storedAt < ROOT_FOLDER_CACHE_MS) {
        cachedRootFolderNames = folderNameSetFromList(stored.names);
        cachedRootFolderNamesAt = Date.now();
        return cachedRootFolderNames;
      }
    } catch {
      // Fall through to a live SharePoint listing.
    }
  }

  const children = await listFolderChildren(getRootFolder());
  const names = new Set();
  for (const item of children) {
    if (!item?.folder) continue;
    const name = String(item.name || '').trim();
    if (!name || name.startsWith('_') || name.startsWith('.')) continue;
    names.add(name.toLowerCase());
  }
  cachedRootFolderNames = names;
  cachedRootFolderNamesAt = Date.now();

  writeStateJson(FOLDER_NAME_CACHE_RELATIVE, {
    updatedAt: new Date().toISOString(),
    names: [...names],
  }).catch((error) => logError('Failed to cache SharePoint project folder names', error));

  return names;
}

function invalidateRootProjectFolderCache() {
  cachedRootFolderNames = null;
  cachedRootFolderNamesAt = 0;
}

async function readStateJson(relativePath) {
  const itemPath = buildStateItemPath(relativePath);
  return withAccessToken(async (accessToken) => {
    try {
      const response = await axios.get(buildItemContentUrl(itemPath), {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
        timeout: getUploadTimeoutMs(),
        responseType: 'text',
        transformResponse: [(data) => data],
        validateStatus: (status) => status >= 200 && status < 300,
      });
      const parsed = JSON.parse(response.data || 'null');
      return parsed && typeof parsed === 'object' ? parsed : null;
    } catch (error) {
      if (isItemNotFoundError(error)) {
        return null;
      }
      throw formatGraphError(error, 'read');
    }
  });
}

async function writeStateJson(relativePath, data) {
  const itemPath = buildStateItemPath(relativePath);
  const body = Buffer.from(JSON.stringify(data, null, 2), 'utf8');
  return withAccessToken(async (accessToken) => {
    let lastError;
    for (let attempt = 1; attempt <= 4; attempt += 1) {
      try {
        await axios.put(buildItemContentUrl(itemPath), body, {
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
          },
          timeout: getUploadTimeoutMs(),
          maxBodyLength: Infinity,
          maxContentLength: Infinity,
          validateStatus: (status) => status >= 200 && status < 300,
        });
        return itemPath;
      } catch (error) {
        lastError = error;
        if (attempt < 4 && (isResourceModifiedError(error) || error?.response?.status === 429)) {
          await delay(200 * attempt);
          continue;
        }
        throw formatGraphError(error, 'write');
      }
    }
    throw formatGraphError(lastError, 'write');
  });
}

async function deleteStateJson(relativePath) {
  return deleteFolderByPath(buildStateItemPath(relativePath));
}

module.exports = {
  uploadToSharePoint,
  uploadFile,
  buildSharePointPath,
  buildProjectFolderPath,
  sanitizeFolderName,
  deleteFolderByPath,
  formatMegabytes,
  clearCachedToken,
  verifySharePointAuth,
  isSharePointConfigError,
  isNameAlreadyExistsError,
  isResourceModifiedError,
  isSharePointConflictSkipError,
  listFolderChildren,
  listFolderFileNames,
  getRootProjectFolderNameSet,
  invalidateRootProjectFolderCache,
  readStateJson,
  writeStateJson,
  deleteStateJson,
  getStateRootFolder,
};
