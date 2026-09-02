const persistentJson = require('./persistentJson.service');

function safeSegment(value) {
  return String(value || 'unknown').replace(/[^a-zA-Z0-9._-]/g, '_');
}

function sessionPath(projectId, documentId) {
  return `in-progress-uploads/${safeSegment(projectId)}/${safeSegment(documentId)}.json`;
}

async function load(projectId, documentId) {
  if (projectId == null || documentId == null) {
    return null;
  }
  const stored = await persistentJson.read(sessionPath(projectId, documentId));
  if (!stored || typeof stored !== 'object' || !stored.uploadUrl) {
    return null;
  }
  return stored;
}

async function save(projectId, documentId, payload) {
  if (projectId == null || documentId == null || !payload?.uploadUrl) {
    return null;
  }
  return persistentJson.write(sessionPath(projectId, documentId), {
    projectId: String(projectId),
    documentId: String(documentId),
    filename: payload.filename || null,
    sharePointPath: payload.sharePointPath || null,
    uploadUrl: payload.uploadUrl,
    totalSize: Number(payload.totalSize) || 0,
    nextOffset: Math.max(0, Number(payload.nextOffset) || 0),
    expirationDateTime: payload.expirationDateTime || null,
    updatedAt: new Date().toISOString(),
  });
}

async function clear(projectId, documentId) {
  if (projectId == null || documentId == null) {
    return;
  }
  await persistentJson.remove(sessionPath(projectId, documentId));
}

module.exports = {
  load,
  save,
  clear,
};
