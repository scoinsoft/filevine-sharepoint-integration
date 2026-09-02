const path = require('path');
const filevineService = require('./filevine.service');
const sharepointService = require('./sharepoint.service');
const syncHistoryService = require('./syncHistory.service');
const projectUploadHistoryService = require('./projectUploadHistory.service');
const { validateSharePointEnv } = require('../config/env');
const { isServerless, shouldStopForDeadline } = require('../config/runtime');
const { log, logError } = require('../utils/logger');

function readPositiveIntEnv(name, fallback) {
  const value = Number(process.env[name]);
  if (!Number.isFinite(value) || value <= 0) {
    return fallback;
  }
  return Math.floor(value);
}

function normalizeFilenameKey(filename) {
  return path.basename(String(filename || '')).trim().toLowerCase();
}

const SYNC_CONCURRENCY = isServerless()
  ? Math.min(readPositiveIntEnv('SYNC_CONCURRENCY', 4), 6)
  : readPositiveIntEnv('SYNC_CONCURRENCY', 6);
const LARGE_FILE_BYTES = 32 * 1024 * 1024;

function isLargeDocument(document) {
  const size = Number(document?.size);
  if (!Number.isFinite(size) || size <= 0) {
    return isServerless();
  }
  return isServerless() && size >= LARGE_FILE_BYTES;
}

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
  if ((options.trigger || 'manual') !== 'scheduled') {
    return null;
  }

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

function finishArchivedProjectSkip({
  projectId,
  projectName,
  projectMeta,
  results,
  options,
  runFolder,
  startedAt,
  emit,
}) {
  const summary = {
    success: true,
    skippedArchivedProject: true,
    projectId,
    projectName,
    projectNumber: projectMeta.projectNumber || null,
    phaseName: projectMeta.phaseName || null,
    isArchived: true,
    total: 0,
    succeeded: 0,
    failed: 0,
    counts: {
      ...buildSyncSummaryCounts(results),
      totalDocuments: 0,
    },
    results,
    message: `Archived: ${projectName}`,
  };
  const history = persistProjectSyncHistory(summary, { ...options, runFolder }, startedAt);
  emit('started', {
    projectId,
    projectName,
    projectNumber: projectMeta.projectNumber || null,
    total: 0,
    skippedArchivedProject: true,
    phaseName: projectMeta.phaseName,
    isArchived: true,
    message: summary.message,
  });
  emit('complete', { ...summary, historyFile: history?.relativePath || null });
  return { ...summary, historyFile: history?.relativePath || null };
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
    validateSharePointEnv();
    emit('status', {
      stage: 'authenticating-sharepoint',
      message: 'Verifying SharePoint connection…',
    });
    await sharepointService.verifySharePointAuth();
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
      stage: 'checking-project',
      message: `Checking project status for ${projectName}…`,
      projectId,
      projectName,
    });

    const projectMeta = await withTokenRetry('getProject', (token) =>
      filevineService.getProject(token, projectId)
    );

    if (projectMeta.isArchived) {
      emit('status', {
        stage: 'skipped-archived',
        message: `Archived: ${projectName}`,
        projectId,
        projectName,
        projectNumber: projectMeta.projectNumber || null,
        phaseName: projectMeta.phaseName,
        isArchived: true,
      });
      return finishArchivedProjectSkip({
        projectId,
        projectName,
        projectMeta,
        results,
        options,
        runFolder,
        startedAt,
        emit,
      });
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
      try {
        await projectUploadHistoryService.markProjectUploaded(projectId, projectName, {
          uploadedCount: 0,
          folderLabel: filevineService.getProjectFolderLabel(projectId, projectName),
        });
      } catch (indexError) {
        logError('Failed to mark empty project in upload history index', indexError);
      }
      const history = persistProjectSyncHistory(
        summary,
        { ...options, runFolder },
        startedAt
      );
      emit('complete', { ...summary, historyFile: history?.relativePath || null });
      return { ...summary, historyFile: history?.relativePath || null };
    }

    const manifest = await filevineService.readProjectUploadManifest(projectId, projectName);
    const uploadedDocumentIds = manifest.uploadedDocumentIds;
    const uploadedFilenames = manifest.uploadedFilenames;
    try {
      const existingSharePointNames = await sharepointService.listFolderFileNames(
        sharepointService.buildProjectFolderPath(projectName)
      );
      for (const name of existingSharePointNames) {
        if (name) uploadedFilenames.add(name);
      }
    } catch (error) {
      logError('Failed to list existing SharePoint files for skip checks', error);
    }
    const uploadedFilenameKeys = new Set(
      [...uploadedFilenames].map((name) => normalizeFilenameKey(name))
    );
    const skipDuplicateDocumentIds = new Set();
    const seenFilenameKeys = new Set(uploadedFilenameKeys);
    for (const doc of documents) {
      const documentKey = String(doc.documentId);
      if (uploadedDocumentIds.has(documentKey)) continue;
      if (!filevineService.hasFileExtension(doc.filename)) continue;
      const filenameKey = normalizeFilenameKey(doc.filename);
      if (!filenameKey) continue;
      if (seenFilenameKeys.has(filenameKey)) {
        skipDuplicateDocumentIds.add(documentKey);
      } else {
        seenFilenameKeys.add(filenameKey);
      }
    }
    if (skipDuplicateDocumentIds.size > 0) {
      log('Duplicate filenames in project; skipping later occurrences', {
        projectId,
        projectName,
        skipCount: skipDuplicateDocumentIds.size,
      });
    }
    const failedHistory = await filevineService.readFailedUploadHistory(projectId, projectName);
    const failedByDocumentId = failedHistory.failedByDocumentId;

    let nextIndex = 0;
    let fatalSharePointError = null;
    let stoppedForDeadline = false;
    const transferProgressThrottle = new Map();
    let largeFileChain = Promise.resolve();

    function withLargeFileGate(document, operation) {
      if (!isLargeDocument(document)) {
        return operation();
      }
      const run = largeFileChain.then(operation, operation);
      largeFileChain = run.then(
        () => undefined,
        () => undefined
      );
      return run;
    }

    function emitTransferProgress(documentId, filename, progress) {
      const bytesUploaded = Number(progress.bytesUploaded) || 0;
      const bytesDownloaded = Number(progress.bytesDownloaded) || 0;
      const bytesTotal = Number(progress.bytesTotal) || 0;
      const stage = progress.stage === 'downloading' ? 'downloading' : 'uploading';
      const bytesDone = stage === 'downloading' ? bytesDownloaded : bytesUploaded;
      const percent =
        typeof progress.percent === 'number'
          ? progress.percent
          : bytesTotal > 0
            ? Math.min(100, Math.round((bytesDone / bytesTotal) * 100))
            : 0;

      const key = String(documentId);
      const now = Date.now();
      const last = transferProgressThrottle.get(key) || { at: 0, percent: -1, stage: null };
      const force =
        percent >= 100 ||
        percent === 0 ||
        last.stage !== stage ||
        Math.abs(percent - last.percent) >= 1;
      if (!force && now - last.at < 500) {
        return;
      }
      transferProgressThrottle.set(key, { at: now, percent, stage });

      emit('file-transfer-progress', {
        documentId: key,
        filename: progress.filename || filename,
        projectId,
        projectName,
        stage,
        bytesUploaded,
        bytesDownloaded,
        bytesTotal,
        bytesDone,
        percent,
        reusedLocalFile: Boolean(progress.reusedLocalFile),
      });
    }

    async function streamUploadWithProgress(downloadUrl, mimeType, filename, documentId, fileSize) {
      return sharepointService.uploadFromDownloadUrl(
        downloadUrl,
        projectId,
        projectName,
        mimeType,
        {
          filename,
          fileSize,
          documentId,
          onProgress: (progress) => {
            emitTransferProgress(documentId, filename, {
              ...progress,
              stage: 'uploading',
              bytesDownloaded: progress.bytesUploaded,
            });
          },
        }
      );
    }

    function skipDuplicateFilename(item) {
      const skippedItem = {
        ...item,
        skippedDuplicateFilename: true,
      };
      results.succeeded.push(skippedItem);
      emit('file-success', {
        ...skippedItem,
        total,
        succeeded: results.succeeded.length,
        failed: results.failed.length,
        message: `Skipped duplicate filename: ${item.filename}`,
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
    }

    async function processDocument(document) {
      if (fatalSharePointError) return;

      if (shouldStopForDeadline()) {
        stoppedForDeadline = true;
        return;
      }

      const item = {
        documentId: document.documentId,
        filename: document.filename,
        folderName: document.folderName,
      };
      const documentKey = String(document.documentId);

      if (!filevineService.hasFileExtension(item.filename)) {
        const skippedItem = {
          ...item,
          skippedNoExtension: true,
        };
        results.succeeded.push(skippedItem);
        emit('file-success', {
          ...skippedItem,
          total,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
          message: `Skipped (no file extension): ${item.filename}`,
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

      if (skipDuplicateDocumentIds.has(documentKey)) {
        skipDuplicateFilename(item);
        filevineService
          .recordProjectUploadSuccess(
            projectId,
            projectName,
            document.documentId,
            item.filename,
            uploadedDocumentIds,
            uploadedFilenames
          )
          .catch((err) => logError('Failed to record skipped duplicate in manifest', err));
        return;
      }

      const hintedSize = Number(document.size) || 0;

      const outcome = await withLargeFileGate(document, async () => {
        if (shouldStopForDeadline()) {
          stoppedForDeadline = true;
          return 'paused';
        }

        emit('progress', {
          stage: 'uploading',
          current: completed + 1,
          total,
          percent: Math.round((completed / total) * 100),
          currentFile: item.filename,
          message: isLargeDocument(document)
            ? `Streaming large file ${item.filename} to SharePoint…`
            : `Streaming ${item.filename} to SharePoint…`,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
        });
        emitTransferProgress(document.documentId, item.filename, {
          stage: 'uploading',
          bytesUploaded: 0,
          bytesDownloaded: 0,
          bytesTotal: hintedSize,
          percent: 0,
        });

        try {
          const source = await withTokenRetry('getDownloadLink', (token) =>
            filevineService.getDocumentDownloadSource(
              token,
              document.documentId,
              document.filename
            )
          );

          const upload = await streamUploadWithProgress(
            source.downloadLink,
            document.contentType || 'application/octet-stream',
            source.filename,
            document.documentId,
            hintedSize
          );

          await filevineService.recordProjectUploadSuccess(
            projectId,
            projectName,
            document.documentId,
            source.filename,
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
            filename: source.filename,
            sharePointPath: upload.sharePointPath,
            size: upload.size || hintedSize || null,
            streamed: true,
          };
          results.succeeded.push(successItem);

          emit('file-success', {
            ...successItem,
            total,
            succeeded: results.succeeded.length,
            failed: results.failed.length,
            message: `Uploaded ${source.filename}`,
          });
          return 'done';
        } catch (error) {
          if (sharepointService.isSharePointConfigError(error)) {
            fatalSharePointError = error;
            throw error;
          }

          if (sharepointService.isSharePointConflictSkipError(error)) {
            const conflictCode = sharepointService.isResourceModifiedError(error)
              ? 'resourceModified'
              : 'nameAlreadyExists';
            log('Skipping SharePoint conflict (already exists / concurrent upload)', {
              projectId,
              documentId: document.documentId,
              filename: item.filename,
              errorCode: conflictCode,
            });

            await filevineService.recordProjectUploadSuccess(
              projectId,
              projectName,
              document.documentId,
              item.filename,
              uploadedDocumentIds,
              uploadedFilenames
            );
            await filevineService.clearFailedUpload(
              projectId,
              projectName,
              document.documentId,
              failedByDocumentId
            );

            const skippedItem = {
              ...item,
              skippedNameConflict: true,
              skippedAlreadyUploaded: true,
              skippedResourceModified: conflictCode === 'resourceModified',
            };
            results.succeeded.push(skippedItem);
            emit('file-success', {
              ...skippedItem,
              total,
              succeeded: results.succeeded.length,
              failed: results.failed.length,
              message: `Skipped (already on SharePoint): ${item.filename}`,
            });
            return 'done';
          }

          if (error?.code === 'UPLOAD_DEADLINE') {
            stoppedForDeadline = true;
            const uploadedMb = sharepointService.formatMegabytes(error.bytesUploaded || 0);
            const totalMb = sharepointService.formatMegabytes(error.totalSize || hintedSize || 0);
            log('Paused large file transfer for function time limit; session saved for resume', {
              projectId,
              documentId: document.documentId,
              filename: item.filename,
              bytesUploaded: error.bytesUploaded || 0,
              totalSize: error.totalSize || hintedSize || 0,
            });
            emit('progress', {
              stage: 'paused',
              current: completed,
              total,
              percent: Math.round((completed / total) * 100),
              currentFile: item.filename,
              message: `Paused ${item.filename} at ${uploadedMb} / ${totalMb} MB; will continue on the next run…`,
              succeeded: results.succeeded.length,
              failed: results.failed.length,
            });
            return 'paused';
          }

          logError(`Sync failed for document ${document.documentId}`, error);

          const failItem = {
            ...item,
            error: error.message,
            errorCode: sharepointService.isNameAlreadyExistsError?.(error)
              ? 'nameAlreadyExists'
              : sharepointService.isResourceModifiedError?.(error)
                ? 'resourceModified'
                : error.code || null,
            localFilePath: null,
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
          return 'done';
        }
      });

      if (outcome === 'paused') {
        return;
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
          if (fatalSharePointError || stoppedForDeadline) return;
          const currentIndex = nextIndex;
          nextIndex += 1;
          if (currentIndex >= documents.length) break;
          try {
            await processDocument(documents[currentIndex]);
          } catch (error) {
            if (sharepointService.isSharePointConfigError(error)) {
              fatalSharePointError = error;
              return;
            }
          }
        }
      })()
    );
    await Promise.all(workers);

    try {
      await filevineService.flushProjectUploadManifest(
        projectId,
        projectName,
        uploadedDocumentIds,
        uploadedFilenames
      );
    } catch (flushError) {
      logError('Failed to flush upload manifest', flushError);
    }

    if (fatalSharePointError) {
      throw fatalSharePointError;
    }

    const counts = buildSyncSummaryCounts(results);
    const processed = results.succeeded.length + results.failed.length;
    const incomplete = Boolean(stoppedForDeadline && processed < total);
    const summary = {
      success: results.failed.length === 0,
      incomplete,
      projectId,
      projectName,
      total,
      succeeded: results.succeeded.length,
      failed: results.failed.length,
      counts: {
        ...counts,
        totalDocuments: total,
      },
      results,
      message: incomplete
        ? `Paused after ${processed} of ${total} file(s) to stay within the host time limit`
        : results.failed.length === 0
          ? `${counts.newlyUploaded} newly uploaded, ${counts.skippedAlreadyUploaded} skipped`
          : `Finished with ${counts.newlyUploaded} newly uploaded, ${counts.skippedAlreadyUploaded} skipped, ${counts.failed} failed`,
    };

    // Always mark the project as known after a completed sync attempt so refresh
    // does not keep showing it as "new" (even if some files failed).
    try {
      await projectUploadHistoryService.markProjectUploaded(projectId, projectName, {
        uploadedCount: uploadedDocumentIds.size,
        folderLabel: filevineService.getProjectFolderLabel(projectId, projectName),
      });
    } catch (indexError) {
      logError('Failed to mark project in upload history index', indexError);
    }

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
    const failurePayload = {
      ...failure,
      historyFile: history?.relativePath || null,
      sharePointConfigError: sharepointService.isSharePointConfigError(error),
    };
    emit('error', failurePayload);
    const wrappedError = error instanceof Error ? error : new Error(String(error));
    wrappedError.syncSummary = failurePayload;
    wrappedError.historyFile = history?.relativePath || null;
    wrappedError.sharePointConfigError = failurePayload.sharePointConfigError;
    throw wrappedError;
  }
}

module.exports = {
  listAllProjects,
  syncProject,
};
