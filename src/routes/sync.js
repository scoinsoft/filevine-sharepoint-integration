const express = require('express');
const filevineService = require('../services/filevine.service');
const sharepointService = require('../services/sharepoint.service');
const { validatePowerAutomateEnv } = require('../config/env');
const { log, logError } = require('../utils/logger');

const router = express.Router();

function sendSse(res, event, data) {
  res.write(`event: ${event}\n`);
  res.write(`data: ${JSON.stringify(data)}\n\n`);
}

function initSse(res) {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();
}

router.get('/projects', async (req, res) => {
  try {
    const offset = Math.max(0, Number(req.query.offset) || 0);
    const limit = Math.max(1, Math.min(1000, Number(req.query.limit) || 1000));
    const accessToken = await filevineService.authenticate();
    const page = await filevineService.listProjectsPage(accessToken, { offset, limit });

    res.json({
      success: true,
      count: page.projects.length,
      offset: page.offset,
      limit: page.limit,
      loadedTo: page.offset + page.projects.length,
      hasMore: page.hasMore,
      projects: page.projects.map(({ projectId, projectName, projectNumber, phaseName, createdDate }) => ({
        projectId,
        projectName,
        projectNumber,
        phaseName,
        createdDate,
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

router.post('/projects/:projectId/sync', async (req, res) => {
  const { projectId } = req.params;
  const projectName =
    req.body?.projectName || req.query?.projectName || `Project ${projectId}`;

  initSse(res);

  const results = {
    succeeded: [],
    failed: [],
  };

  let total = 0;
  let completed = 0;

  try {
    validatePowerAutomateEnv();
    sendSse(res, 'status', {
      stage: 'authenticating',
      message: 'Connecting to Filevine…',
    });

    const accessToken = await filevineService.authenticate();

    sendSse(res, 'status', {
      stage: 'listing',
      message: `Loading documents for ${projectName}…`,
      projectId,
      projectName,
    });

    const documents = await filevineService.listDocuments(accessToken, projectId);
    total = documents.length;

    sendSse(res, 'started', {
      projectId,
      projectName,
      total,
      message: total === 0 ? 'No documents to sync' : `Found ${total} document(s)`,
    });

    if (total === 0) {
      sendSse(res, 'complete', {
        success: true,
        projectId,
        projectName,
        total: 0,
        succeeded: 0,
        failed: 0,
        results,
      });
      return res.end();
    }

    for (const document of documents) {
      const item = {
        documentId: document.documentId,
        filename: document.filename,
        folderName: document.folderName,
      };

      sendSse(res, 'progress', {
        stage: 'downloading',
        current: completed + 1,
        total,
        percent: Math.round((completed / total) * 100),
        currentFile: item.filename,
        message: `Downloading ${item.filename}…`,
        succeeded: results.succeeded.length,
        failed: results.failed.length,
      });

      try {
        const download = await filevineService.downloadDocument(
          accessToken,
          document.documentId,
          document.filename
        );

        sendSse(res, 'progress', {
          stage: 'uploading',
          current: completed + 1,
          total,
          percent: Math.round((completed / total) * 100),
          currentFile: download.filename,
          message: `Uploading ${download.filename} to SharePoint…`,
          succeeded: results.succeeded.length,
          failed: results.failed.length,
        });

        const upload = await sharepointService.uploadFile(
          download.filePath,
          projectId,
          projectName,
          download.mimeType
        );

        const successItem = {
          ...item,
          filename: download.filename,
          sharePointPath: upload.sharePointPath,
          size: download.size,
        };
        results.succeeded.push(successItem);

        sendSse(res, 'file-success', {
          ...successItem,
          message: `Uploaded ${download.filename}`,
        });
      } catch (error) {
        logError(`Sync failed for document ${document.documentId}`, error);

        const failItem = {
          ...item,
          error: error.message,
        };
        results.failed.push(failItem);

        sendSse(res, 'file-error', {
          ...failItem,
          message: `Failed: ${item.filename}`,
        });
      }

      completed += 1;
      sendSse(res, 'progress', {
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

    sendSse(res, 'complete', {
      success: results.failed.length === 0,
      projectId,
      projectName,
      total,
      succeeded: results.succeeded.length,
      failed: results.failed.length,
      results,
      message:
        results.failed.length === 0
          ? `All ${results.succeeded.length} file(s) uploaded successfully`
          : `Finished with ${results.succeeded.length} success, ${results.failed.length} failed`,
    });

    res.end();
  } catch (error) {
    logError('Project sync failed', error);
    sendSse(res, 'error', {
      success: false,
      projectId,
      projectName,
      error: error.message,
      results,
      total,
      completed,
    });
    res.end();
  }
});

router.get('/test-sync', async (req, res) => {
  let projectName;
  let projectId;
  let filename;

  try {
    const accessToken = await filevineService.authenticate();

    ({ projectId, projectName } = await filevineService.getProjects(accessToken));

    const document = await filevineService.getDocuments(accessToken, projectId);
    filename = document.filename;

    const download = await filevineService.downloadDocument(
      accessToken,
      document.documentId,
      document.filename
    );

    validatePowerAutomateEnv();

    try {
      const upload = await sharepointService.uploadFile(
        download.filePath,
        projectId,
        projectName,
        download.mimeType
      );

      const result = {
        success: true,
        projectName,
        filename: download.filename,
        uploaded: true,
        sharePointPath: upload.sharePointPath,
      };

      log('Sync completed', result);
      return res.json(result);
    } catch (uploadError) {
      logError('SharePoint upload failed (sync continued)', uploadError);

      const sharePointPath = sharepointService.buildSharePointPath(
        projectName,
        download.filename
      );

      return res.json({
        success: false,
        projectName,
        filename: download.filename,
        uploaded: false,
        sharePointPath,
        error: uploadError.message,
        downloaded: true,
        savedPath: download.filePath,
      });
    }
  } catch (error) {
    logError('Sync failed', error);
    res.status(500).json({
      success: false,
      projectName: projectName || null,
      filename: filename || null,
      uploaded: false,
      error: error.message,
      details: process.env.NODE_ENV === 'development' ? error.stack : undefined,
    });
  }
});

module.exports = router;
