const fs = require('fs');
const path = require('path');
const { log, logError } = require('../utils/logger');

const UPLOAD_HISTORY_DIR = path.join(process.cwd(), 'upload_history');
const INDEX_FILE = path.join(UPLOAD_HISTORY_DIR, 'projects-index.json');

/** @type {{ version: number, updatedAt: string|null, projects: Record<string, object> } | null} */
let cachedIndex = null;

function emptyIndex() {
  return {
    version: 1,
    updatedAt: null,
    projects: {},
  };
}

function ensureUploadHistoryRoot() {
  if (!fs.existsSync(UPLOAD_HISTORY_DIR)) {
    fs.mkdirSync(UPLOAD_HISTORY_DIR, { recursive: true });
  }
  return UPLOAD_HISTORY_DIR;
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
  ensureUploadHistoryRoot();
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

    // Folder exists under upload_history → treat as already uploaded even without a manifest.
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

function writeIndex(index) {
  ensureUploadHistoryRoot();
  const payload = {
    version: 1,
    updatedAt: new Date().toISOString(),
    projects: index.projects || {},
  };
  fs.writeFileSync(INDEX_FILE, JSON.stringify(payload, null, 2), 'utf8');
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
  ensureUploadHistoryRoot();

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

function loadIndex({ forceRescan = false } = {}) {
  if (cachedIndex && !forceRescan) {
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

function getUploadedProjectIds() {
  const index = loadIndex();
  return Object.keys(index.projects);
}

function getProjectUploadHistorySummary() {
  const index = loadIndex();
  const uploadedProjectIds = Object.keys(index.projects);
  return {
    success: true,
    updatedAt: index.updatedAt,
    uploadedCount: uploadedProjectIds.length,
    uploadedProjectIds,
    projects: uploadedProjectIds.map((projectId) => index.projects[projectId]),
  };
}

function isProjectUploaded(projectId) {
  if (projectId == null) return false;
  const index = loadIndex();
  return Boolean(index.projects[String(projectId)]);
}

function markProjectUploaded(projectId, projectName, options = {}) {
  if (projectId == null) {
    return null;
  }

  try {
    // Always merge from disk so concurrent writers don't overwrite each other.
    const index = readIndexFromDisk();
    mergeScannedIntoIndex(index, scanUploadHistoryFolders());

    const key = String(projectId);
    const now = new Date().toISOString();
    const uploadedCount =
      typeof options.uploadedCount === 'number'
        ? options.uploadedCount
        : Number(options.uploadedCount) || undefined;

    const resolvedCount = typeof uploadedCount === 'number' ? uploadedCount : 0;
    // Index-only by default. Physical upload_history folders are created only when
    // a real successful upload writes uploaded-success.json (or on failures under failed_history).
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

    writeIndex(index);
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

function rebuildProjectUploadHistoryIndex() {
  // Merge folders into the existing index — never wipe index-only entries.
  cachedIndex = null;
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
};
