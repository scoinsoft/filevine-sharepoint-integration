const axios = require('axios');
const fs = require('fs');
const path = require('path');
const { pipeline } = require('stream/promises');
const { filevine } = require('../config/env');
const { log, logError } = require('../utils/logger');
const { inspectDocumentResponse, logInspectionReport } = require('../utils/documentInspector');
const projectUploadHistoryService = require('./projectUploadHistory.service');

const DOWNLOAD_LINK_RESPONSE_FILE = path.join(process.cwd(), 'downloads', 'download-link-response.json');

let cachedToken = null;

function buildAuthHeaders(accessToken) {
  return {
    Authorization: `Bearer ${accessToken}`,
    'x-fv-orgId': filevine.orgId(),
    'x-fv-userId': filevine.userId(),
    'Content-Type': 'application/json',
    Accept: 'application/json',
  };
}

function sanitizeFolderName(name) {
  return String(name || '')
    .replace(/[~"#%&*:<>?/\\{|}]/g, '_')
    .trim() || 'Unnamed Project';
}

function sanitizeFileName(name) {
  const base = path.basename(String(name || 'document'));
  const cleaned = base
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .replace(/[~"#%&*{|}]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
  return cleaned || 'document';
}

function getProjectFolderLabel(projectId, projectName) {
  if (projectId == null) {
    return null;
  }
  return projectName
    ? `${sanitizeFolderName(projectName)}_${projectId}`
    : String(projectId);
}

function ensureDownloadsDir(projectId, projectName) {
  const parts = [process.cwd(), 'downloads'];
  const folderLabel = getProjectFolderLabel(projectId, projectName);
  if (folderLabel) {
    parts.push(folderLabel);
  }

  const downloadsDir = path.join(...parts);
  if (!fs.existsSync(downloadsDir)) {
    fs.mkdirSync(downloadsDir, { recursive: true });
  }
  return downloadsDir;
}

function ensureUploadHistoryDir(projectId, projectName) {
  const folderLabel = getProjectFolderLabel(projectId, projectName);
  if (!folderLabel) {
    throw new Error('projectId is required for upload history path');
  }

  const historyDir = path.join(process.cwd(), 'upload_history', folderLabel);
  if (!fs.existsSync(historyDir)) {
    fs.mkdirSync(historyDir, { recursive: true });
  }
  return historyDir;
}

function ensureFailedHistoryDir(projectId, projectName) {
  const folderLabel = getProjectFolderLabel(projectId, projectName);
  if (!folderLabel) {
    throw new Error('projectId is required for failed history path');
  }

  const historyDir = path.join(process.cwd(), 'failed_history', folderLabel);
  if (!fs.existsSync(historyDir)) {
    fs.mkdirSync(historyDir, { recursive: true });
  }
  return historyDir;
}

function getLegacyUploadManifestPath(projectId, projectName) {
  return path.join(ensureDownloadsDir(projectId, projectName), 'uploaded-success.json');
}

function migrateLegacyUploadManifest(projectId, projectName, manifestPath) {
  const legacyPath = getLegacyUploadManifestPath(projectId, projectName);
  if (fs.existsSync(manifestPath) || !fs.existsSync(legacyPath)) {
    return;
  }

  try {
    fs.mkdirSync(path.dirname(manifestPath), { recursive: true });
    fs.copyFileSync(legacyPath, manifestPath);
    fs.unlinkSync(legacyPath);
    log('Migrated upload manifest to upload_history', { from: legacyPath, to: manifestPath });
  } catch (error) {
    logError('Failed to migrate legacy upload manifest', error);
  }
}

function formatAxiosError(context, error) {
  const status = error.response?.status;
  const data = error.response?.data;
  const message = error.message || 'Unknown error';
  const detail = data ? JSON.stringify(data) : message;
  return new Error(`${context} (status: ${status ?? 'n/a'}): ${detail}`);
}

async function authenticate() {
  const params = new URLSearchParams({
    client_id: filevine.clientId(),
    client_secret: filevine.clientSecret(),
    grant_type: 'personal_access_token',
    scope: 'fv.api.gateway.access tenant filevine.v2.api.* openid email fv.auth.tenant.read',
    token: filevine.pat(),
  });

  try {
    const response = await axios.post(filevine.tokenUrl(), params.toString(), {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    });

    cachedToken = response.data.access_token;
    if (!cachedToken) {
      throw new Error('Token response did not include access_token');
    }

    log('Access Token acquired');
    return cachedToken;
  } catch (error) {
    logError('Filevine authentication failed', error);
    throw formatAxiosError('Filevine authentication failed', error);
  }
}

function getApiClient(accessToken) {
  return axios.create({
    baseURL: filevine.apiBase(),
    headers: buildAuthHeaders(accessToken),
  });
}

function normalizeProject(project) {
  const projectId = project.projectId?.native ?? project.projectId;
  const projectName = project.projectName || project.name || 'Unknown Project';
  return {
    projectId,
    projectName,
    projectNumber: project.projectNumber ?? project.number ?? null,
    phaseName: project.phaseName || project.phase?.name || null,
    createdDate: project.createdDate || project.createDate || null,
    isArchived: project.isArchived === true,
    raw: project,
  };
}

function hasFileExtension(filename) {
  const base = path.basename(String(filename || '').trim());
  const ext = path.extname(base);
  return ext.length > 1;
}

function normalizeDocument(document) {
  const documentId = document.documentId?.native ?? document.documentId;
  const filename = document.filename || document.fileName || document.name || 'document';
  const folderName = document.folderName || document.folder?.name || null;
  return {
    documentId,
    filename,
    folderName,
    size: document.size ?? document.fileSize ?? null,
    contentType: document.contentType || document.mimeType || null,
    uploadedDate: document.uploadedDate || document.createDate || null,
    raw: document,
  };
}

async function fetchAllPages(client, endpoint, label) {
  const collected = [];
  let offset = 0;
  const limit = 50;
  let hasMore = true;

  while (hasMore) {
    const response = await client.get(endpoint, {
      params: { offset, limit },
    });

    const items = response.data?.items || [];
    collected.push(...items);

    const count = response.data?.count;
    const nextOffset = response.data?.offset ?? offset;
    const hasMoreFlag = response.data?.hasMore;

    log(`${label} page`, {
      offset,
      received: items.length,
      totalCollected: collected.length,
      count,
      hasMore: hasMoreFlag,
    });

    if (items.length === 0) {
      hasMore = false;
    } else if (typeof hasMoreFlag === 'boolean') {
      hasMore = hasMoreFlag;
      offset = nextOffset + items.length;
    } else if (typeof count === 'number') {
      offset += items.length;
      hasMore = offset < count;
    } else if (items.length < limit) {
      hasMore = false;
    } else {
      offset += items.length;
    }

    // Safety guard against unexpected infinite loops
    if (offset > 10000) {
      log(`Stopping ${label} pagination early (offset > 10000)`);
      break;
    }
  }

  return collected;
}

async function getProject(accessToken, projectId) {
  const client = getApiClient(accessToken);

  try {
    const response = await client.get(`/projects/${projectId}`);
    return normalizeProject(response.data || {});
  } catch (error) {
    logError('Failed to get project', error);
    throw formatAxiosError(`Failed to get project ${projectId}`, error);
  }
}

async function listProjectsPage(accessToken, options = {}) {
  const client = getApiClient(accessToken);
  const offset = Math.max(0, Number(options.offset) || 0);
  const limit = Math.max(1, Math.min(1000, Number(options.limit) || 1000));

  try {
    const response = await client.get('/projects', {
      params: { offset, limit },
    });

    const items = response.data?.items || [];
    const hasMoreFlag = response.data?.hasMore;
    const raw = response.data || {};
    const rawCount = Number(raw.count);
    const candidateTotal = Number(raw.total ?? raw.totalCount ?? raw.totalItems ?? raw.recordCount);
    let total = null;

    if (Number.isFinite(candidateTotal) && candidateTotal >= 0) {
      total = candidateTotal;
    } else if (
      Number.isFinite(rawCount) &&
      rawCount >= 0 &&
      (hasMoreFlag === false || rawCount > offset + items.length)
    ) {
      // Some APIs return only page size in `count`; only trust it when it behaves like a true total.
      total = rawCount;
    }

    const projects = items
      .map(normalizeProject)
      .filter((project) => project.projectId != null);

    let hasMore;
    if (typeof hasMoreFlag === 'boolean') {
      hasMore = hasMoreFlag;
    } else if (typeof total === 'number') {
      hasMore = offset + projects.length < total;
    } else {
      // If total is unknown, a full page strongly suggests there may be more.
      hasMore = items.length >= limit;
    }

    return {
      projects,
      offset,
      limit,
      total,
      hasMore,
    };
  } catch (error) {
    logError('Failed to list projects page', error);
    throw formatAxiosError('Failed to list projects page', error);
  }
}

async function listDocuments(accessToken, projectId) {
  const client = getApiClient(accessToken);

  try {
    const items = await fetchAllPages(
      client,
      `/projects/${projectId}/documents`,
      `Documents for project ${projectId}`
    );
    log(`Documents count: ${items.length}`);

    return items
      .map(normalizeDocument)
      .filter((document) => document.documentId != null);
  } catch (error) {
    logError('Failed to list documents', error);
    throw formatAxiosError(`Failed to list documents for project ${projectId}`, error);
  }
}

async function getDownloadLink(accessToken, documentId) {
  const client = getApiClient(accessToken);

  try {
    const response = await client.post('/Documents/batch/download', {
      DocumentIds: [documentId],
      DownloadUrlTimeToLive: 600,
    });

    log('Batch download link response', response.data);

    try {
      ensureDownloadsDir();
      fs.writeFileSync(DOWNLOAD_LINK_RESPONSE_FILE, JSON.stringify(response.data, null, 2), 'utf8');
      log(`Saved batch download response to ${DOWNLOAD_LINK_RESPONSE_FILE}`);
    } catch (error) {
      if (error.code === 'ENOSPC') {
        logError('Skipped saving batch download debug file (disk full)', error);
      } else {
        throw error;
      }
    }

    const items = Array.isArray(response.data) ? response.data : [];
    if (items.length === 0) {
      throw new Error('Batch download response did not return any items');
    }

    const first = items[0];
    const downloadLink = first?.downloadLink;

    if (!downloadLink || typeof downloadLink !== 'string') {
      throw new Error(
        `Batch download response missing downloadLink. Keys: ${Object.keys(first || {}).join(', ')}`
      );
    }

    log('Download URL', downloadLink);
    console.log('✓ Download URL acquired');
    console.log('Download Link:', downloadLink);

    return {
      downloadLink,
      batchResponse: response.data,
      savedTo: DOWNLOAD_LINK_RESPONSE_FILE,
      versionKey: first?.versionKey || null,
    };
  } catch (error) {
    logError('Failed to get download link', error);
    throw formatAxiosError(`Failed to get download link for document ${documentId}`, error);
  }
}

function getLocalDownloadPath(filename, projectId, projectName, documentId) {
  const safeName = sanitizeFileName(filename);
  const projectDir = ensureDownloadsDir(projectId, projectName);
  // Isolate each document under its own folder to avoid same-name races
  // when SYNC_CONCURRENCY downloads/uploads in parallel.
  const targetDir =
    documentId != null ? path.join(projectDir, String(documentId)) : projectDir;
  if (!fs.existsSync(targetDir)) {
    fs.mkdirSync(targetDir, { recursive: true });
  }
  return {
    safeName,
    filePath: path.join(targetDir, safeName),
  };
}

function getProjectUploadManifestPath(projectId, projectName, { create = false } = {}) {
  const folderLabel = getProjectFolderLabel(projectId, projectName);
  if (!folderLabel) {
    throw new Error('projectId is required for upload history path');
  }
  const historyDir = path.join(process.cwd(), 'upload_history', folderLabel);
  if (create && !fs.existsSync(historyDir)) {
    fs.mkdirSync(historyDir, { recursive: true });
  }
  return path.join(historyDir, 'uploaded-success.json');
}

function readProjectUploadManifest(projectId, projectName) {
  const manifestPath = getProjectUploadManifestPath(projectId, projectName, { create: false });
  if (fs.existsSync(manifestPath)) {
    // no-op: already present
  } else {
    migrateLegacyUploadManifest(projectId, projectName, manifestPath);
  }

  if (!fs.existsSync(manifestPath)) {
    return {
      manifestPath,
      uploadedDocumentIds: new Set(),
      uploadedFilenames: new Set(),
    };
  }

  try {
    const raw = fs.readFileSync(manifestPath, 'utf8');
    const parsed = JSON.parse(raw);
    const ids = Array.isArray(parsed?.uploadedDocumentIds) ? parsed.uploadedDocumentIds : [];
    const names = Array.isArray(parsed?.uploadedFilenames) ? parsed.uploadedFilenames : [];
    return {
      manifestPath,
      uploadedDocumentIds: new Set(ids.map((id) => String(id))),
      uploadedFilenames: new Set(names.map((name) => String(name))),
    };
  } catch (error) {
    logError('Failed to read upload manifest (starting fresh)', error);
    return {
      manifestPath,
      uploadedDocumentIds: new Set(),
      uploadedFilenames: new Set(),
    };
  }
}

let manifestWriteChain = Promise.resolve();
const failedHistoryWriteChains = new Map();

function getFailedHistoryPath(projectId, projectName, { create = false } = {}) {
  const folderLabel = getProjectFolderLabel(projectId, projectName);
  if (!folderLabel) {
    throw new Error('projectId is required for failed history path');
  }
  const historyDir = path.join(process.cwd(), 'failed_history', folderLabel);
  if (create && !fs.existsSync(historyDir)) {
    fs.mkdirSync(historyDir, { recursive: true });
  }
  return path.join(historyDir, 'failed-history.json');
}

function readFailedUploadHistory(projectId, projectName) {
  // Do not create folders just by reading — only when recording a real failure.
  const historyPath = getFailedHistoryPath(projectId, projectName, { create: false });
  if (!fs.existsSync(historyPath)) {
    return {
      historyPath,
      failedByDocumentId: new Map(),
    };
  }

  try {
    const raw = fs.readFileSync(historyPath, 'utf8');
    const parsed = JSON.parse(raw);
    const entries = Array.isArray(parsed?.failed) ? parsed.failed : [];
    const failedByDocumentId = new Map();
    for (const entry of entries) {
      const key = String(entry?.documentId ?? '');
      if (!key) continue;
      failedByDocumentId.set(key, {
        documentId: key,
        filename: entry.filename || null,
        folderName: entry.folderName || null,
        error: entry.error || 'Unknown error',
        errorCode: entry.errorCode || null,
        firstFailedAt: entry.firstFailedAt || new Date().toISOString(),
        lastFailedAt: entry.lastFailedAt || new Date().toISOString(),
      });
    }
    return { historyPath, failedByDocumentId };
  } catch (error) {
    logError('Failed to read failed history (starting fresh)', error);
    return {
      historyPath,
      failedByDocumentId: new Map(),
    };
  }
}

function queueFailedHistoryWrite(projectId, projectName, writer) {
  const key = getProjectFolderLabel(projectId, projectName) || String(projectId);
  const chain = failedHistoryWriteChains.get(key) || Promise.resolve();

  const next = chain
    .then(() => writer())
    .catch((error) => {
      logError('Failed to update failed history', error);
      return null;
    });

  failedHistoryWriteChains.set(
    key,
    next.finally(() => {
      if (failedHistoryWriteChains.get(key) === next) {
        failedHistoryWriteChains.delete(key);
      }
    })
  );

  return next;
}

function removeFailedHistoryIfEmpty(projectId, projectName) {
  const historyPath = getFailedHistoryPath(projectId, projectName, { create: false });
  const historyDir = path.dirname(historyPath);
  try {
    if (fs.existsSync(historyPath)) {
      fs.unlinkSync(historyPath);
    }
    if (fs.existsSync(historyDir)) {
      const remaining = fs.readdirSync(historyDir);
      if (remaining.length === 0) {
        fs.rmdirSync(historyDir);
      }
    }
  } catch (error) {
    logError('Failed to clean empty failed history folder', error);
  }
}

function writeFailedHistoryFile(projectId, projectName, failedByDocumentId) {
  if (!failedByDocumentId || failedByDocumentId.size === 0) {
    removeFailedHistoryIfEmpty(projectId, projectName);
    return null;
  }

  const historyPath = getFailedHistoryPath(projectId, projectName, { create: true });
  const payload = {
    projectId: String(projectId),
    projectName,
    updatedAt: new Date().toISOString(),
    failed: [...failedByDocumentId.values()].sort((a, b) =>
      String(a.documentId).localeCompare(String(b.documentId))
    ),
  };
  fs.writeFileSync(historyPath, JSON.stringify(payload, null, 2), 'utf8');
  return historyPath;
}

function extractFailureCode(errorMessage) {
  const message = String(errorMessage || '');
  const codeMatch = message.match(/"code"\s*:\s*"([^"]+)"/);
  if (codeMatch) return codeMatch[1];
  if (message.includes('ENOENT')) return 'ENOENT';
  if (message.includes('socket hang up')) return 'socket_hang_up';
  if (message.includes('ECONNRESET')) return 'ECONNRESET';
  if (message.includes('resourceModified')) return 'resourceModified';
  if (message.includes('nameAlreadyExists')) return 'nameAlreadyExists';
  if (message.includes('Failed to read upload chunk')) return 'chunk_read_mismatch';
  return null;
}

function recordFailedUpload(projectId, projectName, failItem, failedByDocumentId) {
  const documentKey = String(failItem?.documentId ?? '');
  if (!documentKey) {
    return Promise.resolve(null);
  }

  const existing = failedByDocumentId.get(documentKey);
  const now = new Date().toISOString();
  const error = failItem?.error || 'Unknown error';
  failedByDocumentId.set(documentKey, {
    documentId: documentKey,
    filename: failItem?.filename || existing?.filename || null,
    folderName: failItem?.folderName || existing?.folderName || null,
    error,
    errorCode: failItem?.errorCode || extractFailureCode(error) || existing?.errorCode || null,
    localFilePath: failItem?.localFilePath || existing?.localFilePath || null,
    firstFailedAt: existing?.firstFailedAt || now,
    lastFailedAt: now,
  });

  log('Recording failed upload', {
    projectId,
    projectName,
    documentId: documentKey,
    filename: failItem?.filename || null,
    errorCode: failedByDocumentId.get(documentKey).errorCode,
    error,
  });

  return queueFailedHistoryWrite(projectId, projectName, () =>
    writeFailedHistoryFile(projectId, projectName, failedByDocumentId)
  );
}

function clearFailedUpload(projectId, projectName, documentId, failedByDocumentId) {
  const documentKey = String(documentId ?? '');
  if (!documentKey) {
    return Promise.resolve(null);
  }

  if (!failedByDocumentId.has(documentKey)) {
    return Promise.resolve(null);
  }

  failedByDocumentId.delete(documentKey);
  return queueFailedHistoryWrite(projectId, projectName, () =>
    writeFailedHistoryFile(projectId, projectName, failedByDocumentId)
  );
}

function recordProjectUploadSuccess(
  projectId,
  projectName,
  documentId,
  filename,
  uploadedDocumentIds,
  uploadedFilenames
) {
  manifestWriteChain = manifestWriteChain
    .then(() => {
      // Only create the upload_history folder when we actually have a successful upload.
      const manifestPath = getProjectUploadManifestPath(projectId, projectName, { create: true });
      const documentKey = String(documentId);
      const fileKey = String(filename || '');

      uploadedDocumentIds.add(documentKey);
      if (fileKey) {
        uploadedFilenames.add(fileKey);
      }

      const payload = {
        projectId: String(projectId),
        projectName,
        updatedAt: new Date().toISOString(),
        uploadedDocumentIds: [...uploadedDocumentIds],
        uploadedFilenames: [...uploadedFilenames],
      };

      fs.writeFileSync(manifestPath, JSON.stringify(payload, null, 2), 'utf8');
      try {
        projectUploadHistoryService.markProjectUploaded(projectId, projectName, {
          uploadedCount: uploadedDocumentIds.size,
          folderLabel: getProjectFolderLabel(projectId, projectName),
        });
      } catch (indexError) {
        logError('Failed to update project upload history index', indexError);
      }
      return manifestPath;
    })
    .catch((error) => {
      logError('Failed to write upload manifest (upload still counted in this run)', error);
      return null;
    });

  return manifestWriteChain;
}

function getExistingDownload(filename, projectId, projectName, documentId) {
  const { safeName, filePath } = getLocalDownloadPath(
    filename,
    projectId,
    projectName,
    documentId
  );
  if (!fs.existsSync(filePath)) {
    return null;
  }

  const stats = fs.statSync(filePath);
  if (!stats.isFile() || stats.size <= 0) {
    return null;
  }

  return {
    filePath,
    filename: safeName,
    size: stats.size,
    mimeType: 'application/octet-stream',
    alreadyDownloaded: true,
  };
}

function isRetryableDownloadError(error) {
  const code = error?.code || error?.cause?.code;
  if (['ECONNRESET', 'ETIMEDOUT', 'ECONNABORTED', 'ENOTFOUND', 'EAI_AGAIN', 'ENOENT'].includes(code)) {
    return true;
  }
  const message = String(error?.message || '');
  return (
    message.includes('socket hang up') ||
    message.includes('ECONNRESET') ||
    message.includes('ETIMEDOUT') ||
    message.includes('ENOENT')
  );
}

function delayMs(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function downloadFromPresignedUrl(downloadLink, filename, projectId, projectName, options = {}) {
  const documentId = options.documentId;
  const { safeName, filePath } = getLocalDownloadPath(
    filename,
    projectId,
    projectName,
    documentId
  );
  const onProgress = typeof options.onProgress === 'function' ? options.onProgress : null;
  const maxAttempts = 3;
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const tempPath = `${filePath}.partial`;
    try {
      fs.mkdirSync(path.dirname(filePath), { recursive: true });
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
      }

      const response = await axios.get(downloadLink, {
        responseType: 'stream',
        timeout: 10 * 60 * 1000,
      });

      const totalBytes = Number(response.headers['content-length']) || 0;
      let bytesDownloaded = 0;
      let lastReportedAt = 0;

      if (onProgress) {
        onProgress({
          filename: safeName,
          bytesDownloaded: 0,
          bytesTotal: totalBytes,
          percent: 0,
          stage: 'downloading',
        });
      }

      response.data.on('data', (chunk) => {
        bytesDownloaded += chunk.length;
        if (!onProgress) return;

        const now = Date.now();
        const isComplete = totalBytes > 0 && bytesDownloaded >= totalBytes;
        if (!isComplete && now - lastReportedAt < 200) return;
        lastReportedAt = now;

        const percent =
          totalBytes > 0 ? Math.min(100, Math.round((bytesDownloaded / totalBytes) * 100)) : 0;
        onProgress({
          filename: safeName,
          bytesDownloaded,
          bytesTotal: totalBytes,
          percent,
          stage: 'downloading',
        });
      });

      await pipeline(response.data, fs.createWriteStream(tempPath));
      fs.renameSync(tempPath, filePath);

      const stats = fs.statSync(filePath);
      const mimeType = response.headers['content-type'] || 'unknown';

      if (onProgress) {
        onProgress({
          filename: safeName,
          bytesDownloaded: stats.size,
          bytesTotal: totalBytes || stats.size,
          percent: 100,
          stage: 'downloading',
        });
      }

      log('Downloaded successfully', {
        filename: safeName,
        size: stats.size,
        mimeType,
        savedPath: filePath,
        documentId: documentId || null,
        attempt,
      });

      return {
        filePath,
        filename: safeName,
        size: stats.size,
        mimeType,
        alreadyDownloaded: false,
      };
    } catch (error) {
      lastError = error;
      try {
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
      } catch {
        // ignore cleanup errors
      }

      logError('File download failed', {
        filename: safeName,
        documentId: documentId || null,
        attempt,
        maxAttempts,
        code: error.code || null,
        message: error.message,
        filePath,
      });

      if (attempt < maxAttempts && isRetryableDownloadError(error)) {
        await delayMs(1000 * attempt);
        continue;
      }
      throw formatAxiosError(`Failed to download file from presigned URL`, error);
    }
  }

  throw formatAxiosError(`Failed to download file from presigned URL`, lastError);
}

async function inspectDocumentMetadata(accessToken, documentId) {
  const client = getApiClient(accessToken);

  try {
    const response = await client.get(`/documents/${documentId}`);
    const metadata = response.data;

    log('Document metadata response (complete)', metadata);

    const report = inspectDocumentResponse(metadata);
    logInspectionReport(report);

    return report;
  } catch (error) {
    logError('Failed to inspect document metadata', error);
    throw formatAxiosError(`Failed to inspect document ${documentId}`, error);
  }
}

async function downloadDocument(accessToken, documentId, filename, projectId, projectName, options = {}) {
  const resolvedFilename = filename || 'document';
  const onProgress = typeof options.onProgress === 'function' ? options.onProgress : null;

  const existing = getExistingDownload(resolvedFilename, projectId, projectName, documentId);
  if (existing) {
    log('Skipping download; file already exists locally', {
      filename: existing.filename,
      savedPath: existing.filePath,
      size: existing.size,
      projectId,
      documentId,
    });
    if (onProgress) {
      onProgress({
        filename: existing.filename,
        bytesDownloaded: existing.size,
        bytesTotal: existing.size,
        percent: 100,
        stage: 'downloading',
        reusedLocalFile: true,
      });
    }
    return {
      ...existing,
      documentId,
      downloadLink: null,
      downloadLinkResponseFile: null,
      batchResponse: null,
    };
  }

  const { downloadLink, savedTo, versionKey, batchResponse } = await getDownloadLink(
    accessToken,
    documentId
  );

  const downloadFilename = filename || versionKey || 'document';
  const file = await downloadFromPresignedUrl(
    downloadLink,
    downloadFilename,
    projectId,
    projectName,
    { onProgress, documentId }
  );

  return {
    ...file,
    documentId,
    downloadLink,
    downloadLinkResponseFile: savedTo,
    batchResponse,
  };
}

function deleteLocalDownload(filePath) {
  if (!filePath || !fs.existsSync(filePath)) {
    return false;
  }

  fs.unlinkSync(filePath);
  log('Deleted local download after successful upload', { filePath });
  return true;
}

function clearCachedToken() {
  cachedToken = null;
}

module.exports = {
  authenticate,
  clearCachedToken,
  getProjectFolderLabel,
  listProjectsPage,
  getProject,
  listDocuments,
  getDownloadLink,
  downloadFromPresignedUrl,
  downloadDocument,
  deleteLocalDownload,
  readProjectUploadManifest,
  recordProjectUploadSuccess,
  readFailedUploadHistory,
  recordFailedUpload,
  clearFailedUpload,
  inspectDocumentMetadata,
  hasFileExtension,
};
