const fs = require('fs');
const path = require('path');
const filevineService = require('./filevine.service');
const { log, logError } = require('../utils/logger');

const SYNC_HISTORY_DIR = path.join(process.cwd(), 'sync_history');
const RUN_SUMMARY_FILENAME = 'run-summary.json';

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function toRelativePath(filePath) {
  return path.relative(process.cwd(), filePath).split(path.sep).join('/');
}

function formatRunFolderName(isoString) {
  const date = new Date(isoString);
  const pad = (value, length = 2) => String(value).padStart(length, '0');
  return [
    `${date.getUTCFullYear()}-${pad(date.getUTCMonth() + 1)}-${pad(date.getUTCDate())}`,
    `${pad(date.getUTCHours())}-${pad(date.getUTCMinutes())}-${pad(date.getUTCSeconds())}-${pad(date.getUTCMilliseconds(), 3)}`,
  ].join('_');
}

function createSyncRunFolder(startedAt, trigger = 'manual') {
  const folderName = formatRunFolderName(startedAt);
  const dirPath = path.join(SYNC_HISTORY_DIR, folderName);
  ensureDir(dirPath);

  const runFolder = {
    folderName,
    dirPath,
    relativeDir: toRelativePath(dirPath),
    startedAt,
    trigger,
  };

  log('Created sync run history folder', {
    trigger,
    folder: runFolder.relativeDir,
  });

  return runFolder;
}

function getProjectHistoryFilename(projectId, projectName) {
  const folderLabel = filevineService.getProjectFolderLabel(projectId, projectName);
  if (!folderLabel) {
    throw new Error('projectId is required for sync history filename');
  }
  return `${folderLabel}.json`;
}

function writeJsonFile(dirPath, filename, record) {
  const filePath = path.join(dirPath, filename);
  fs.writeFileSync(filePath, JSON.stringify(record, null, 2), 'utf8');
  return {
    filePath,
    relativePath: toRelativePath(filePath),
  };
}

function categorizeSyncResults(results = {}) {
  const skippedAlreadyUploaded = [];
  const newlyUploaded = [];
  const failed = [];

  for (const item of results.succeeded || []) {
    const entry = {
      documentId: item.documentId,
      filename: item.filename,
      folderName: item.folderName || null,
    };

    if (item.skippedAlreadyUploaded || item.skippedNoExtension) {
      skippedAlreadyUploaded.push({
        ...entry,
        reason: item.skippedNoExtension ? 'no_extension' : 'already_uploaded',
      });
      continue;
    }

    newlyUploaded.push({
      ...entry,
      sharePointPath: item.sharePointPath || null,
      size: item.size ?? null,
      reusedLocalFile: Boolean(item.reusedLocalFile),
    });
  }

  for (const item of results.failed || []) {
    failed.push({
      documentId: item.documentId,
      filename: item.filename,
      folderName: item.folderName || null,
      error: item.error || 'Unknown error',
      reason: item.skippedTooLarge ? 'file_too_large' : 'upload_error',
    });
  }

  return {
    skippedAlreadyUploaded,
    newlyUploaded,
    failed,
  };
}

function buildCounts(categorized, totalDocuments = 0) {
  const failedTooLarge = categorized.failed.filter((item) => item.reason === 'file_too_large').length;
  return {
    totalDocuments,
    newlyUploaded: categorized.newlyUploaded.length,
    skippedAlreadyUploaded: categorized.skippedAlreadyUploaded.length,
    failed: categorized.failed.length,
    failedTooLarge,
    failedUpload: categorized.failed.length - failedTooLarge,
  };
}

function buildProjectSyncRecord(summary, meta = {}) {
  const categorized = categorizeSyncResults(summary.results);
  const finishedAt = meta.finishedAt || new Date().toISOString();
  const startedAt = meta.startedAt || finishedAt;
  const durationMs =
    typeof meta.durationMs === 'number'
      ? meta.durationMs
      : Math.max(0, new Date(finishedAt).getTime() - new Date(startedAt).getTime());

  return {
    version: 1,
    trigger: meta.trigger || 'manual',
    scheduledRunId: meta.scheduledRunId || null,
    runFolder: meta.runFolder?.relativeDir || null,
    projectId: String(summary.projectId),
    projectName: summary.projectName,
    startedAt,
    finishedAt,
    durationMs,
    success: summary.success !== false && categorized.failed.length === 0 && !summary.error,
    skippedArchivedProject: Boolean(summary.skippedArchivedProject),
    projectNumber: summary.projectNumber || null,
    phaseName: summary.phaseName || null,
    error: summary.error || null,
    counts: buildCounts(categorized, summary.total ?? 0),
    newlyUploaded: categorized.newlyUploaded,
    skippedAlreadyUploaded: categorized.skippedAlreadyUploaded,
    failed: categorized.failed,
  };
}

function saveProjectSyncHistory(summary, meta = {}) {
  try {
    if (!meta.runFolder?.dirPath) {
      throw new Error('runFolder is required to save project sync history');
    }

    const finishedAt = meta.finishedAt || new Date().toISOString();
    const record = buildProjectSyncRecord(summary, { ...meta, finishedAt });
    const filename = getProjectHistoryFilename(summary.projectId, summary.projectName);
    const saved = writeJsonFile(meta.runFolder.dirPath, filename, record);

    log('Saved project sync history', {
      projectId: summary.projectId,
      projectName: summary.projectName,
      trigger: record.trigger,
      runFolder: meta.runFolder.relativeDir,
      file: saved.relativePath,
      counts: record.counts,
    });

    return {
      record,
      runFolder: meta.runFolder,
      ...saved,
    };
  } catch (error) {
    logError('Failed to save project sync history', error);
    return null;
  }
}

function buildScheduledRunRecord(runMeta = {}, projectEntries = []) {
  const finishedAt = runMeta.finishedAt || new Date().toISOString();
  const startedAt = runMeta.startedAt || finishedAt;
  const durationMs =
    typeof runMeta.durationMs === 'number'
      ? runMeta.durationMs
      : Math.max(0, new Date(finishedAt).getTime() - new Date(startedAt).getTime());

  const aggregateCounts = {
    totalDocuments: 0,
    newlyUploaded: 0,
    skippedAlreadyUploaded: 0,
    failed: 0,
    failedTooLarge: 0,
    failedUpload: 0,
    projectsSucceeded: 0,
    projectsFailed: 0,
    projectsSkippedArchived: 0,
  };

  const projects = projectEntries.map((entry) => {
    const counts = entry.record?.counts || {
      totalDocuments: 0,
      newlyUploaded: 0,
      skippedAlreadyUploaded: 0,
      failed: 0,
      failedTooLarge: 0,
      failedUpload: 0,
    };

    aggregateCounts.totalDocuments += counts.totalDocuments;
    aggregateCounts.newlyUploaded += counts.newlyUploaded;
    aggregateCounts.skippedAlreadyUploaded += counts.skippedAlreadyUploaded;
    aggregateCounts.failed += counts.failed;
    aggregateCounts.failedTooLarge += counts.failedTooLarge;
    aggregateCounts.failedUpload += counts.failedUpload;

    if (entry.record?.skippedArchivedProject) {
      aggregateCounts.projectsSkippedArchived += 1;
      aggregateCounts.projectsSucceeded += 1;
    } else if (entry.record?.success) {
      aggregateCounts.projectsSucceeded += 1;
    } else {
      aggregateCounts.projectsFailed += 1;
    }

    return {
      projectId: entry.projectId,
      projectName: entry.projectName,
      success: Boolean(entry.record?.success),
      skippedArchivedProject: Boolean(entry.record?.skippedArchivedProject),
      error: entry.error || entry.record?.error || null,
      counts,
      historyFile: entry.historyFile || null,
    };
  });

  return {
    version: 1,
    trigger: 'scheduled',
    runId: runMeta.runId || startedAt,
    runFolder: runMeta.runFolder?.relativeDir || null,
    startedAt,
    finishedAt,
    durationMs,
    projectCount: projects.length,
    success: aggregateCounts.projectsFailed === 0 && aggregateCounts.failed === 0,
    counts: aggregateCounts,
    projects,
  };
}

function saveScheduledRunHistory(runMeta = {}, projectEntries = []) {
  try {
    if (!runMeta.runFolder?.dirPath) {
      throw new Error('runFolder is required to save scheduled sync run history');
    }

    const finishedAt = runMeta.finishedAt || new Date().toISOString();
    const record = buildScheduledRunRecord({ ...runMeta, finishedAt }, projectEntries);
    const saved = writeJsonFile(runMeta.runFolder.dirPath, RUN_SUMMARY_FILENAME, record);

    log('Saved scheduled sync run history', {
      runId: record.runId,
      runFolder: runMeta.runFolder.relativeDir,
      file: saved.relativePath,
      counts: record.counts,
    });

    return {
      record,
      runFolder: runMeta.runFolder,
      ...saved,
    };
  } catch (error) {
    logError('Failed to save scheduled sync run history', error);
    return null;
  }
}

module.exports = {
  RUN_SUMMARY_FILENAME,
  buildCounts,
  buildProjectSyncRecord,
  buildScheduledRunRecord,
  categorizeSyncResults,
  createSyncRunFolder,
  saveProjectSyncHistory,
  saveScheduledRunHistory,
};
