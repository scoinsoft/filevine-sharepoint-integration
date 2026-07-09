const path = require('path');
const filevineService = require('./filevine.service');
const sharepointService = require('./sharepoint.service');
const syncHistoryService = require('./syncHistory.service');
const { validatePowerAutomateEnv } = require('../config/env');
const { log, logError } = require('../utils/logger');
const { Semaphore } = require('../utils/semaphore');

function readPositiveIntEnv(name, fallback) {
  const value = Number(process.env[name]);
  if (!Number.isFinite(value) || value <= 0) {
    return fallback;
  }
  return Math.floor(value);
}

const SYNC_CONCURRENCY = readPositiveIntEnv('SYNC_CONCURRENCY', 6);
const SYNC_LARGE_FILE_MB = readPositiveIntEnv('SYNC_LARGE_FILE_MB', 25);
const SYNC_LARGE_FILE_CONCURRENCY = readPositiveIntEnv('SYNC_LARGE_FILE_CONCURRENCY', 1);
const LARGE_FILE_BYTES = SYNC_LARGE_FILE_MB * 1024 * 1024;

function isUnauthorizedError(error) {
  const status = error?.response?.status;
  if (status === 401) return true;
  const message = String(error?.message || '');
  return message.includes('status: 401') || message.includes('status code 401');
}

async function listAllProjects() {
  const accessToken = await filevineService.authenticate();
  const projects = [];
  let offset = 0;

  while (true) {
    const page = await filevineService.listProjectsPage(accessToken, { offset, limit: 1000 });
    projects.push(...page.projects);
    if (!page.hasMore) break;
    offset += 1000;
  }

  return projects;
}

function buildSyncSummaryCounts(results) {
  const categorized = syncHistoryService.categorizeSyncResults(results);
  return syncHistoryService.buildCounts(categorized, 0);
}

function persistProjectSyncHistory(summary, options, startedAt) {
  const finishedAt = new Date().toISOString();
  return syncHistoryService.saveProjectSyncHistory(summary, {
    trigger: options.trigger || 'manual',
    scheduledRunId: options.scheduledRunId || null,
    runFolder: options.runFolder,
    startedAt,
    finishedAt,
    durationMs: Math.max(0, new Date(finishedAt).getTime() - new Date(startedAt).getTime()),
  });
}

async function syncProject(projectId, projectName, options = {}) {
  const onEvent = typeof options.onEvent === 'function' ? options.onEvent : () => {};
  const emit = (event, data) => onEvent(event, data);
  const startedAt = new Date().toISOString();
  const runFolder =
    options.runFolder ||
    syncHistoryService.createSyncRunFolder(startedAt, options.trigger || 'manual');

  const results = {
    succeeded: [],
    failed: [],
  };

  let total = 0;
  let completed = 0;

  try {
    validatePowerAutomateEnv();
    emit('status', {
      stage: 'authenticating',
      message: 'Connecting to Filevine…',
    });

    let accessToken = await filevineService.authenticate();
    let refreshingTokenPromise = null;

    async function refreshAccessToken() {
      if (!refreshingTokenPromise) {
        refreshingTokenPromise = (async () => {
          filevineService.clearCachedToken();
          const fresh = await filevineService.authenticate();
          log('Refreshed Filevine access token after 401', { projectId, projectName });
          return fresh;
        })().finally(() => {
          refreshingTokenPromise = null;
        });
      }
      accessToken = await refreshingTokenPromise;
      return accessToken;
    }

    async function withTokenRetry(operationName, operation) {
      try {
        return await operation(accessToken);
      } catch (error) {
        if (!isUnauthorizedError(error)) {
          throw error;
        }
        log(`${operationName} received 401, refreshing token and retrying`, { projectId, projectName });
        const freshToken = await refreshAccessToken();
        return operation(freshToken);
      }
    }

    emit('status', {
      stage: 'listing',
      message: `Loading documents for ${projectName}…`,
      projectId,
      projectName,
    });

    const documents = await withTokenRetry('listDocuments', (token) =>
      filevineService.listDocuments(token, projectId)
    );
    total = documents.length;

    emit('started', {
      projectId,
      projectName,
      total,
      message: total === 0 ? 'No documents to sync' : `Found ${total} document(s)`,
    });

    if (total === 0) {
      const summary = {
        success: true,
        projectId,
        projectName,
        total: 0,
        succeeded: 0,
        failed: 0,
        counts: {
          ...buildSyncSummaryCounts(results),
          totalDocuments: 0,
        },
        results,
      };
      const history = persistProjectSyncHistory(
        summary,
        { ...options, runFolder },
        startedAt
      );
      emit('complete', { ...summary, historyFile: history?.relativePath || null });
      return { ...summary, historyFile: history?.relativePath || null };
    }

    const manifest = filevineService.readProjectUploadManifest(projectId, projectName);
    const uploadedDocumentIds = manifest.uploadedDocumentIds;
    const uploadedFilenames = manifest.uploadedFilenames;
    const failedHistory = filevineService.readFailedUploadHistory(projectId, projectName);
    const failedByDocumentId = failedHistory.failedByDocumentId;

    let nextIndex = 0;
    const largeUploadGate = new Semaphore(SYNC_LARGE_FILE_CONCURRENCY);

    async function uploadWithMemoryGuard(filePath, size, mimeType) {
      const isLarge = Number(size) > LARGE_FILE_BYTES;
      const runUpload = () =>
        sharepointService.uploadFile(filePath, projectId, projectName, mimeType);

      if (!isLarge) {
        return runUpload();
      }

      log('Queueing large file upload', {
        projectId,
        filename: path.basename(filePath),
        size,
        largeFileMb: SYNC_LARGE_FILE_MB,
        maxConcurrentLargeUploads: SYNC_LARGE_FILE_CONCURRENCY,
      });
      return largeUploadGate.run(runUpload);
    }

    async function processDocument(document) {
      const item = {
        documentId: document.documentId,
        filename: document.filename,
        folderName: document.folderName,
      };
      const documentKey = String(document.documentId);

      if (uploadedDocumentIds.has(documentKey)) {
        const skippedItem = {
          ...item,
          skippedAlreadyUploaded: true,
        };
        results.succeeded.push(skippedItem);
        emit('file-success', {
          ...skippedItem,
          total,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
          message: `Skipped already uploaded: ${item.filename}`,
        });

        completed += 1;
        emit('progress', {
          stage: 'file-complete',
          current: completed,
          total,
          percent: Math.round((completed / total) * 100),
          currentFile: item.filename,
          message: `Processed ${completed} of ${total}`,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
        });
        return;
      }

      if (document.size && !sharepointService.isUploadableFileSize(document.size)) {
        const failItem = {
          ...item,
          error:
            `File too large for base64 upload (${sharepointService.formatMegabytes(document.size)} MB). ` +
            `Maximum supported is ${Math.round(sharepointService.getMaxUploadBytes() / (1024 * 1024))} MB.`,
          skippedTooLarge: true,
        };
        results.failed.push(failItem);
        emit('file-error', {
          ...failItem,
          total,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
          message: `Skipped too large: ${item.filename}`,
        });

        completed += 1;
        emit('progress', {
          stage: 'file-complete',
          current: completed,
          total,
          percent: Math.round((completed / total) * 100),
          currentFile: item.filename,
          message: `Processed ${completed} of ${total}`,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
        });
        return;
      }

      let localFilePath = null;

      emit('progress', {
        stage: 'downloading',
        current: completed + 1,
        total,
        percent: Math.round((completed / total) * 100),
        currentFile: item.filename,
        message: `Preparing ${item.filename}…`,
        succeeded: results.succeeded.length,
        failed: results.failed.length,
      });

      try {
        const download = await withTokenRetry('downloadDocument', (token) =>
          filevineService.downloadDocument(
            token,
            document.documentId,
            document.filename,
            projectId,
            projectName
          )
        );
        localFilePath = download.filePath;

        if (!sharepointService.isUploadableFileSize(download.size)) {
          filevineService.deleteLocalDownload(download.filePath);
          localFilePath = null;
          throw new Error(
            `File too large for base64 upload (${sharepointService.formatMegabytes(download.size)} MB). ` +
              `Maximum supported is ${Math.round(sharepointService.getMaxUploadBytes() / (1024 * 1024))} MB.`
          );
        }

        emit('progress', {
          stage: 'uploading',
          current: completed + 1,
          total,
          percent: Math.round((completed / total) * 100),
          currentFile: download.filename,
          message: download.alreadyDownloaded
            ? `Using local copy of ${download.filename}; uploading to SharePoint…`
            : `Uploading ${download.filename} to SharePoint…`,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
        });

        const upload = await uploadWithMemoryGuard(
          download.filePath,
          download.size,
          download.mimeType
        );

        filevineService.deleteLocalDownload(download.filePath);
        localFilePath = null;

        await filevineService.recordProjectUploadSuccess(
          projectId,
          projectName,
          document.documentId,
          download.filename,
          uploadedDocumentIds,
          uploadedFilenames
        );
        await filevineService.clearFailedUpload(
          projectId,
          projectName,
          document.documentId,
          failedByDocumentId
        );

        const successItem = {
          ...item,
          filename: download.filename,
          sharePointPath: upload.sharePointPath,
          size: download.size,
          reusedLocalFile: Boolean(download.alreadyDownloaded),
        };
        results.succeeded.push(successItem);

        emit('file-success', {
          ...successItem,
          total,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
          message: `Uploaded ${download.filename}`,
        });
      } catch (error) {
        logError(`Sync failed for document ${document.documentId}`, error);

        if (localFilePath) {
          log('Keeping local download after failed upload', { filePath: localFilePath });
        }

        const failItem = {
          ...item,
          error: error.message,
        };
        results.failed.push(failItem);
        await filevineService.recordFailedUpload(
          projectId,
          projectName,
          failItem,
          failedByDocumentId
        );

        emit('file-error', {
          ...failItem,
          total,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
          message: `Failed: ${item.filename}`,
        });
      }

      completed += 1;
      emit('progress', {
        stage: 'file-complete',
        current: completed,
        total,
        percent: Math.round((completed / total) * 100),
        currentFile: item.filename,
        message: `Processed ${completed} of ${total}`,
        succeeded: results.succeeded.length,
        failed: results.failed.length,
      });
    }

    const workerCount = Math.min(SYNC_CONCURRENCY, total);
    const workers = Array.from({ length: workerCount }, () =>
      (async () => {
        while (true) {
          const currentIndex = nextIndex;
          nextIndex += 1;
          if (currentIndex >= documents.length) break;
          await processDocument(documents[currentIndex]);
        }
      })()
    );
    await Promise.all(workers);

    const counts = buildSyncSummaryCounts(results);
    const summary = {
      success: results.failed.length === 0,
      projectId,
      projectName,
      total,
      succeeded: results.succeeded.length,
      failed: results.failed.length,
      counts: {
        ...counts,
        totalDocuments: total,
      },      results,
      message:
        results.failed.length === 0
          ? `${counts.newlyUploaded} newly uploaded, ${counts.skippedAlreadyUploaded} skipped (already uploaded)`
          : `Finished with ${counts.newlyUploaded} newly uploaded, ${counts.skippedAlreadyUploaded} skipped, ${counts.failed} failed`,
    };
    const history = persistProjectSyncHistory(summary, { ...options, runFolder }, startedAt);
    emit('complete', { ...summary, historyFile: history?.relativePath || null });
    return { ...summary, historyFile: history?.relativePath || null };
  } catch (error) {
    logError('Project sync failed', error);
    const failure = {
      success: false,
      projectId,
      projectName,
      error: error.message,
      results,
      total,
      completed,
      counts: buildSyncSummaryCounts(results),
    };
    const history = persistProjectSyncHistory(
      {
        ...failure,
        counts: {
          ...buildSyncSummaryCounts(results),
          totalDocuments: total,
        },
      },
      { ...options, runFolder },
      startedAt
    );
    const failurePayload = { ...failure, historyFile: history?.relativePath || null };
    emit('error', failurePayload);
    const wrappedError = error instanceof Error ? error : new Error(String(error));
    wrappedError.syncSummary = failurePayload;
    wrappedError.historyFile = history?.relativePath || null;
    throw wrappedError;
  }
}

module.exports = {
  listAllProjects,
  syncProject,
};
