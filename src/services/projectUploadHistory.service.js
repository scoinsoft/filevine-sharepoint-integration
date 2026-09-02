const fs = require('fs');
const path = require('path');
const { isServerless } = require('../config/runtime');
const { uploadHistoryDir, ensureDir } = require('../config/paths');
const persistentJson = require('./persistentJson.service');
const { log, logError } = require('../utils/logger');

const UPLOAD_HISTORY_DIR = uploadHistoryDir();
const INDEX_FILE = path.join(UPLOAD_HISTORY_DIR, 'projects-index.json');
const INDEX_RELATIVE = 'upload_history/projects-index.json';

/** @type {{ version: number, updatedAt: string|null, projects: Record<string, object> } | null} */
let cachedIndex = null;

function emptyIndex() {
  return {
    version: 1,
    updatedAt: null,
    projects: {},
  };
}

function extractProjectIdFromFolderName(folderName) {
  const match = String(folderName || '').match(/_(\d+)$/);
  return match ? match[1] : null;
}

function sanitizeFolderName(name) {
  return String(name || '')
    .replace(/[~"#%&*:<>?/\\{|}]/g, '_')
    .trim() || 'Unnamed Project';
}

function buildFolderLabel(projectId, projectName, folderLabel) {
  if (folderLabel) return folderLabel;
  if (projectId == null) return null;
  return projectName
    ? `${sanitizeFolderName(projectName)}_${projectId}`
    : String(projectId);
}

function resolveFolderLabel(projectId, projectName, folderLabel) {
  return buildFolderLabel(projectId, projectName, folderLabel);
}

function readManifestInFolder(folderPath) {
  const manifestPath = path.join(folderPath, 'uploaded-success.json');
  if (!fs.existsSync(manifestPath)) {
    return null;
  }
  try {
    const parsed = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
    if (!parsed || typeof parsed !== 'object') {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

function buildProjectEntryFromManifest(manifest, folderLabel) {
  const projectId = String(manifest.projectId || extractProjectIdFromFolderName(folderLabel) || '');
  if (!projectId) {
    return null;
  }

  const uploadedCount = Array.isArray(manifest.uploadedDocumentIds)
    ? manifest.uploadedDocumentIds.length
    : Array.isArray(manifest.uploadedFilenames)
      ? manifest.uploadedFilenames.length
      : 0;

  return {
    projectId,
    projectName: manifest.projectName || folderLabel,
    folderLabel,
    firstSyncedAt: manifest.updatedAt || new Date().toISOString(),
    lastSyncedAt: manifest.updatedAt || new Date().toISOString(),
    uploadedCount,
  };
}

function scanUploadHistoryFolders() {
  if (isServerless() || !fs.existsSync(UPLOAD_HISTORY_DIR)) {
    return {};
  }

  const projects = {};
  let entries = [];
  try {
    entries = fs.readdirSync(UPLOAD_HISTORY_DIR, { withFileTypes: true });
  } catch (error) {
    logError('Failed to scan upload_history folders', error);
    return projects;
  }

  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (entry.name.startsWith('_') || entry.name.startsWith('.')) continue;

    const folderPath = path.join(UPLOAD_HISTORY_DIR, entry.name);
    const manifest = readManifestInFolder(folderPath);
    if (manifest) {
      const projectEntry = buildProjectEntryFromManifest(manifest, entry.name);
      if (projectEntry) {
        projects[projectEntry.projectId] = projectEntry;
        continue;
      }
    }

    const projectId = extractProjectIdFromFolderName(entry.name);
    if (!projectId) continue;
    const nameFromFolder = entry.name.replace(new RegExp(`_${projectId}$`), '') || entry.name;
    let mtime = null;
    try {
      mtime = fs.statSync(folderPath).mtime.toISOString();
    } catch {
      mtime = new Date().toISOString();
    }
    projects[projectId] = {
      projectId,
      projectName: nameFromFolder,
      folderLabel: entry.name,
      firstSyncedAt: mtime,
      lastSyncedAt: mtime,
      uploadedCount: 0,
    };
  }

  return projects;
}

async function writeIndex(index) {
  const payload = {
    version: 1,
    updatedAt: new Date().toISOString(),
    projects: index.projects || {},
  };

  if (isServerless()) {
    await persistentJson.write(INDEX_RELATIVE, payload);
  } else {
    ensureDir(UPLOAD_HISTORY_DIR);
    fs.writeFileSync(INDEX_FILE, JSON.stringify(payload, null, 2), 'utf8');
  }

  cachedIndex = payload;
  return payload;
}

function mergeScannedIntoIndex(index, scanned) {
  let changed = false;

  for (const [projectId, entry] of Object.entries(scanned)) {
    if (!index.projects[projectId]) {
      index.projects[projectId] = entry;
      changed = true;
      continue;
    }

    const existing = index.projects[projectId];
    const nextCount = Math.max(Number(existing.uploadedCount) || 0, Number(entry.uploadedCount) || 0);
    if (nextCount !== (Number(existing.uploadedCount) || 0)) {
      existing.uploadedCount = nextCount;
      changed = true;
    }
    if (!existing.projectName && entry.projectName) {
      existing.projectName = entry.projectName;
      changed = true;
    }
    if (!existing.folderLabel && entry.folderLabel) {
      existing.folderLabel = entry.folderLabel;
      changed = true;
    }
    if (entry.lastSyncedAt && entry.lastSyncedAt > (existing.lastSyncedAt || '')) {
      existing.lastSyncedAt = entry.lastSyncedAt;
      changed = true;
    }
    if (!existing.firstSyncedAt && entry.firstSyncedAt) {
      existing.firstSyncedAt = entry.firstSyncedAt;
      changed = true;
    }
  }

  return changed;
}

function readIndexFromDisk() {
  let index = emptyIndex();
  if (!fs.existsSync(INDEX_FILE)) {
    return index;
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(INDEX_FILE, 'utf8'));
    if (parsed && typeof parsed === 'object') {
      index = {
        version: 1,
        updatedAt: parsed.updatedAt || null,
        projects:
          parsed.projects && typeof parsed.projects === 'object' ? parsed.projects : {},
      };
    }
  } catch (error) {
    logError('Failed to read projects-index.json; rebuilding from folders', error);
  }

  return index;
}

async function loadIndex({ forceRescan = false } = {}) {
  if (cachedIndex && !forceRescan) {
    return cachedIndex;
  }

  if (isServerless()) {
    const stored = await persistentJson.read(INDEX_RELATIVE);
    cachedIndex = stored && stored.projects ? {
      version: 1,
      updatedAt: stored.updatedAt || null,
      projects: stored.projects,
    } : emptyIndex();
    return cachedIndex;
  }

  const index = readIndexFromDisk();
  const scanned = scanUploadHistoryFolders();
  const changed = mergeScannedIntoIndex(index, scanned);

  if (!fs.existsSync(INDEX_FILE) || changed || forceRescan) {
    return writeIndex(index);
  }

  cachedIndex = index;
  return index;
}

async function getUploadedProjectIds() {
  const index = await loadIndex();
  return Object.keys(index.projects);
}

async function getProjectUploadHistorySummary() {
  const index = await loadIndex();
  const uploadedProjectIds = Object.keys(index.projects);
  return {
    success: true,
    updatedAt: index.updatedAt,
    uploadedCount: uploadedProjectIds.length,
    uploadedProjectIds,
    projects: uploadedProjectIds.map((projectId) => index.projects[projectId]),
  };
}

async function isProjectUploaded(projectId) {
  if (projectId == null) return false;
  const index = await loadIndex();
  return Boolean(index.projects[String(projectId)]);
}

async function markProjectUploaded(projectId, projectName, options = {}) {
  if (projectId == null) {
    return null;
  }

  try {
    let index;
    if (isServerless()) {
      const stored = await persistentJson.read(INDEX_RELATIVE);
      index = stored && stored.projects ? {
        version: 1,
        updatedAt: stored.updatedAt || null,
        projects: stored.projects,
      } : emptyIndex();
    } else {
      index = readIndexFromDisk();
      mergeScannedIntoIndex(index, scanUploadHistoryFolders());
    }

    const key = String(projectId);
    const now = new Date().toISOString();
    const uploadedCount =
      typeof options.uploadedCount === 'number'
        ? options.uploadedCount
        : Number(options.uploadedCount) || undefined;

    const resolvedCount = typeof uploadedCount === 'number' ? uploadedCount : 0;
    const folderLabel = resolveFolderLabel(key, projectName, options.folderLabel || null);

    const existing = index.projects[key];
    if (existing) {
      existing.projectName = projectName || existing.projectName;
      existing.lastSyncedAt = now;
      existing.uploadedCount = Math.max(Number(existing.uploadedCount) || 0, resolvedCount);
      if (folderLabel) {
        existing.folderLabel = folderLabel;
      }
    } else {
      index.projects[key] = {
        projectId: key,
        projectName: projectName || `Project ${key}`,
        folderLabel: folderLabel || null,
        firstSyncedAt: now,
        lastSyncedAt: now,
        uploadedCount: resolvedCount,
      };
    }

    await writeIndex(index);
    log('Marked project in upload history index', {
      projectId: key,
      projectName: projectName || null,
      uploadedCount: index.projects[key].uploadedCount,
      folderLabel: index.projects[key].folderLabel,
    });
    return index.projects[key];
  } catch (error) {
    logError('Failed to mark project in upload history index', error);
    return null;
  }
}

async function rebuildProjectUploadHistoryIndex() {
  cachedIndex = null;
  if (isServerless()) {
    return loadIndex({ forceRescan: true });
  }
  const index = readIndexFromDisk();
  mergeScannedIntoIndex(index, scanUploadHistoryFolders());
  return writeIndex(index);
}

module.exports = {
  getProjectUploadHistorySummary,
  getUploadedProjectIds,
  isProjectUploaded,
  markProjectUploaded,
  loadIndex,
  rebuildProjectUploadHistoryIndex,
  sanitizeFolderName,
};
