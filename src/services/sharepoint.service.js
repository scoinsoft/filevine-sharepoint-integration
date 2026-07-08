const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { powerAutomate } = require('../config/env');
const { log, logError } = require('../utils/logger');

function sanitizeFolderName(name) {
  return name.replace(/[~"#%&*:<>?/\\{|}]/g, '_').trim() || 'Unnamed Project';
}

function buildSharePointPath(projectName, filename) {
  const safeProjectName = sanitizeFolderName(projectName);
  return `Documents/Filevine/${safeProjectName}/${filename}`;
}

async function uploadToSharePoint({ projectId, projectName, filename, contentType, buffer }) {
  const uploadUrl = powerAutomate.uploadUrl();
  const sharePointPath = buildSharePointPath(projectName, filename);
  const fileContent = buffer.toString('base64');

  // Power Automate HTTP trigger schema expects projectId as Integer (Express params are strings).
  const numericProjectId = Number(projectId);
  if (!Number.isInteger(numericProjectId)) {
    throw new Error(`Invalid projectId for Power Automate upload: ${projectId}`);
  }

  const payload = {
    projectId: numericProjectId,
    projectName,
    filename,
    contentType,
    fileContent,
  };

  log('Uploading to SharePoint via Power Automate', {
    projectId,
    filename,
    sharePointPath,
    contentType,
    size: buffer.length,
  });

  try {
    const response = await axios.post(uploadUrl, payload, {
      headers: { 'Content-Type': 'application/json' },
      validateStatus: () => true,
    });

    if (response.status !== 200) {
      const detail = typeof response.data === 'string'
        ? response.data
        : JSON.stringify(response.data);
      throw new Error(
        `Power Automate upload failed (status ${response.status}): ${detail}`
      );
    }

    const body = response.data || {};
    if (body.success === false) {
      throw new Error(body.message || 'Power Automate reported upload failure');
    }

    console.log('Upload successful');
    log('Upload successful', {
      projectId,
      filename,
      sharePointPath,
      message: body.message || 'File uploaded successfully',
    });

    return {
      success: true,
      message: body.message || 'File uploaded successfully',
      sharePointPath,
      filename,
      projectId,
      projectName,
    };
  } catch (error) {
    logError('Power Automate upload failed', error);
    log('Upload debug context', { projectId, filename, sharePointPath });

    if (error.response) {
      const detail = typeof error.response.data === 'string'
        ? error.response.data
        : JSON.stringify(error.response.data);
      throw new Error(
        `Power Automate upload failed (status ${error.response.status}): ${detail}`
      );
    }

    throw error;
  }
}

async function uploadFile(filePath, projectId, projectName, contentType) {
  const buffer = fs.readFileSync(filePath);
  const filename = path.basename(filePath);

  return uploadToSharePoint({
    projectId,
    projectName,
    filename,
    contentType: contentType || 'application/octet-stream',
    buffer,
  });
}

module.exports = {
  uploadToSharePoint,
  uploadFile,
  buildSharePointPath,
};
