const axios = require('axios');
const fs = require('fs');
const path = require('path');
const { pipeline } = require('stream/promises');
const { filevine } = require('../config/env');
const { log, logError } = require('../utils/logger');
const { inspectDocumentResponse, logInspectionReport } = require('../utils/documentInspector');

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

function ensureDownloadsDir() {
  const downloadsDir = path.join(process.cwd(), 'downloads');
  if (!fs.existsSync(downloadsDir)) {
    fs.mkdirSync(downloadsDir, { recursive: true });
  }
  return downloadsDir;
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
    raw: project,
  };
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

async function listProjects(accessToken) {
  const client = getApiClient(accessToken);

  try {
    const items = await fetchAllPages(client, '/projects', 'Projects');
    log(`Projects count: ${items.length}`);

    const projects = items
      .map(normalizeProject)
      .filter((project) => project.projectId != null);

    return projects;
  } catch (error) {
    logError('Failed to list projects', error);
    throw formatAxiosError('Failed to list projects', error);
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

async function getProjects(accessToken) {
  try {
    const projects = await listProjects(accessToken);

    if (projects.length === 0) {
      throw new Error('No projects returned from Filevine');
    }

    const project = projects[0];
    log('Project selected', {
      projectName: project.projectName,
      projectId: project.projectId,
    });
    console.log('Project Name:', project.projectName);
    console.log('Project ID:', project.projectId);

    return {
      projectId: project.projectId,
      projectName: project.projectName,
      raw: project.raw,
    };
  } catch (error) {
    if (error.message?.includes('No projects')) {
      throw error;
    }
    throw error;
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

async function getDocuments(accessToken, projectId) {
  try {
    const documents = await listDocuments(accessToken, projectId);

    if (documents.length === 0) {
      throw new Error(`No documents found for project ${projectId}`);
    }

    const document = documents[0];
    log('Document selected', {
      filename: document.filename,
      documentId: document.documentId,
      folderName: document.folderName,
    });
    console.log('Filename:', document.filename);
    console.log('DocumentId:', document.documentId);

    return {
      documentId: document.documentId,
      filename: document.filename,
      folderName: document.folderName,
      raw: document.raw,
    };
  } catch (error) {
    if (error.message?.includes('No documents')) {
      throw error;
    }
    throw error;
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

    ensureDownloadsDir();
    fs.writeFileSync(DOWNLOAD_LINK_RESPONSE_FILE, JSON.stringify(response.data, null, 2), 'utf8');
    log(`Saved batch download response to ${DOWNLOAD_LINK_RESPONSE_FILE}`);

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

async function downloadFromPresignedUrl(downloadLink, filename) {
  const safeName = path.basename(filename);
  const filePath = path.join(ensureDownloadsDir(), safeName);

  try {
    const response = await axios.get(downloadLink, {
      responseType: 'stream',
    });

    await pipeline(response.data, fs.createWriteStream(filePath));

    const stats = fs.statSync(filePath);
    const mimeType = response.headers['content-type'] || 'unknown';

    console.log('✓ Download succeeded');
    console.log('Filename:', safeName);
    console.log('Size:', stats.size, 'bytes');
    console.log('MIME type:', mimeType);
    console.log('Saved path:', filePath);

    log('Downloaded successfully', {
      filename: safeName,
      size: stats.size,
      mimeType,
      savedPath: filePath,
    });

    return {
      filePath,
      filename: safeName,
      size: stats.size,
      mimeType,
    };
  } catch (error) {
    logError('File download failed', error);
    throw formatAxiosError(`Failed to download file from presigned URL`, error);
  }
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

async function downloadDocument(accessToken, documentId, filename) {
  const { downloadLink, savedTo, versionKey, batchResponse } = await getDownloadLink(
    accessToken,
    documentId
  );

  const resolvedFilename = filename || versionKey || 'document';
  const file = await downloadFromPresignedUrl(downloadLink, resolvedFilename);

  return {
    ...file,
    documentId,
    downloadLink,
    downloadLinkResponseFile: savedTo,
    batchResponse,
  };
}

module.exports = {
  authenticate,
  listProjectsPage,
  listProjects,
  getProjects,
  listDocuments,
  getDocuments,
  getDownloadLink,
  downloadFromPresignedUrl,
  downloadDocument,
  inspectDocumentMetadata,
};
