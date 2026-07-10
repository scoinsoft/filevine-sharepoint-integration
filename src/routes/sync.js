const express = require('express');
const filevineService = require('../services/filevine.service');
const syncProjectService = require('../services/syncProject.service');
const scheduleService = require('../services/schedule.service');
const syncRunService = require('../services/syncRun.service');
const projectUploadHistoryService = require('../services/projectUploadHistory.service');
const removeArchivedSharePointFoldersService = require('../services/removeArchivedSharePointFolders.service');
const { log, logError } = require('../utils/logger');

const router = express.Router();

function createSseSender(res) {
  let clientClosed = false;
  res.on('close', () => {
    clientClosed = true;
  });

  return (event, data) => {
    if (clientClosed) return;
    try {
      res.write(`event: ${event}\n`);
      res.write(`data: ${JSON.stringify(data)}\n\n`);
    } catch {
      clientClosed = true;
    }
  };
}

function initSse(res) {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();
}

router.get('/projects/upload-history', async (req, res) => {
  try {
    const rebuild = String(req.query.rebuild || '') === '1';
    if (rebuild) {
      projectUploadHistoryService.rebuildProjectUploadHistoryIndex();
    }
    const summary = projectUploadHistoryService.getProjectUploadHistorySummary();
    res.json(summary);
  } catch (error) {
    logError('Failed to load project upload history', error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

router.get('/projects', async (req, res) => {
  try {
    const offset = Math.max(0, Number(req.query.offset) || 0);
    const limit = Math.max(1, Math.min(1000, Number(req.query.limit) || 1000));
    const accessToken = await filevineService.authenticate();
    const page = await filevineService.listProjectsPage(accessToken, { offset, limit });
    const uploadedIds = new Set(projectUploadHistoryService.getUploadedProjectIds());

    res.json({
      success: true,
      count: page.projects.length,
      offset: page.offset,
      limit: page.limit,
      loadedTo: page.offset + page.projects.length,
      hasMore: page.hasMore,
      uploadedProjectCount: uploadedIds.size,
      projects: page.projects.map(({ projectId, projectName, projectNumber, phaseName, createdDate, isArchived }) => ({
        projectId,
        projectName,
        projectNumber,
        phaseName,
        createdDate,
        isArchived,
        alreadyUploaded: uploadedIds.has(String(projectId)),
        isNew: !uploadedIds.has(String(projectId)),
      })),
    });
  } catch (error) {
    logError('Failed to list projects', error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

router.get('/projects/:projectId/documents', async (req, res) => {
  try {
    const { projectId } = req.params;
    const accessToken = await filevineService.authenticate();
    const documents = await filevineService.listDocuments(accessToken, projectId);

    res.json({
      success: true,
      projectId,
      count: documents.length,
      documents: documents.map(
        ({ documentId, filename, folderName, size, contentType, uploadedDate }) => ({
          documentId,
          filename,
          folderName,
          size,
          contentType,
          uploadedDate,
        })
      ),
    });
  } catch (error) {
    logError('Failed to list documents', error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

router.get('/sync/status', (req, res) => {
  res.json({
    success: true,
    activeRuns: syncRunService.getActiveRuns(),
    hasActiveRuns: syncRunService.hasActiveRuns(),
  });
});

router.get('/sharepoint/remove-archived-folders/preview', async (req, res) => {
  try {
    const preview = await removeArchivedSharePointFoldersService.countArchivedProjects();
    res.json({
      success: true,
      archivedCount: preview.archivedCount,
      projects: preview.projects,
    });
  } catch (error) {
    logError('Failed to preview archived SharePoint folder cleanup', error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
});

router.post('/sharepoint/remove-archived-folders', async (req, res) => {
  if (syncRunService.hasActiveRuns()) {
    return res.status(423).json({
      success: false,
      error: 'Cannot remove folders while a project sync is running. Wait for sync to finish and try again.',
    });
  }

  initSse(res);
  const sendSse = createSseSender(res);

  try {
    log('Archived SharePoint folder cleanup started');
    await removeArchivedSharePointFoldersService.removeArchivedSharePointFolders({
      onEvent: (event, data) => sendSse(event, data),
    });
    log('Archived SharePoint folder cleanup finished');
    res.end();
  } catch (error) {
    logError('Archived SharePoint folder cleanup failed', error);
    if (!res.headersSent) {
      res.status(500).json({
        success: false,
        error: error.message,
      });
      return;
    }
    sendSse('error', { error: error.message });
    sendSse('complete', {
      success: false,
      total: 0,
      deleted: 0,
      skipped: 0,
      failed: 0,
      message: error.message,
    });
    res.end();
  }
});

router.post('/projects/:projectId/sync', async (req, res) => {
  const { projectId } = req.params;
  const projectName =
    req.body?.projectName || req.query?.projectName || `Project ${projectId}`;

  if (scheduleService.isUploadBlocked()) {
    return res.status(423).json({
      success: false,
      error:
        'Uploads are disabled while the scheduled sync is running. This usually takes about 2–3 hours.',
    });
  }

  initSse(res);
  const sendSse = createSseSender(res);
  syncRunService.startRun(projectId, projectName);

  try {
    log('Manual project sync started', { projectId, projectName });
    await syncProjectService.syncProject(projectId, projectName, {
      onEvent: (event, data) => sendSse(event, data),
    });
    log('Manual project sync finished', { projectId, projectName });
    res.end();
  } catch (error) {
    logError('Project sync failed', error);
    if (!res.headersSent) {
      res.status(500).json({
        success: false,
        error: error.message,
      });
      return;
    }
    res.end();
  } finally {
    syncRunService.endRun(projectId);
  }
});

module.exports = router;
