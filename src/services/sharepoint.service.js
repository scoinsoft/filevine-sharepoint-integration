const fs = require('fs');
const path = require('path');
const axios = require('axios');
const { powerAutomate } = require('../config/env');
const { log, logError } = require('../utils/logger');

const DEFAULT_TIMEOUT_MS = 180000;
const DEFAULT_MAX_UPLOAD_MB = 100;
const MAX_RETRIES = 3;

function sanitizeFolderName(name) {
  return name.replace(/[~"#%&*:<>?/\\{|}]/g, '_').trim() || 'Unnamed Project';
}

function buildSharePointPath(projectName, filename) {
  const safeProjectName = sanitizeFolderName(projectName);
  return `Documents/Filevine/${safeProjectName}/${filename}`;
}

function getUploadTimeoutMs() {
  const ms = Number(process.env.POWER_AUTOMATE_TIMEOUT_MS || DEFAULT_TIMEOUT_MS);
  if (!Number.isFinite(ms) || ms <= 0) {
    return DEFAULT_TIMEOUT_MS;
  }
  return Math.floor(ms);
}

function getMaxUploadBytes() {
  const mb = Number(process.env.POWER_AUTOMATE_MAX_UPLOAD_MB || DEFAULT_MAX_UPLOAD_MB);
  if (!Number.isFinite(mb) || mb <= 0) {
    return DEFAULT_MAX_UPLOAD_MB * 1024 * 1024;
  }
  return Math.floor(mb * 1024 * 1024);
}

function formatMegabytes(bytes) {
  return (Number(bytes) / (1024 * 1024)).toFixed(1);
}

function buildFileTooLargeError(filename, sizeBytes) {
  const maxBytes = getMaxUploadBytes();
  const maxMb = Math.round(maxBytes / (1024 * 1024));
  return new Error(
    `File too large for base64 upload: ${filename} (${formatMegabytes(sizeBytes)} MB). ` +
      `Maximum supported is ${maxMb} MB with current upload mode.`
  );
}

function isUploadableFileSize(sizeBytes) {
  const size = Number(sizeBytes);
  if (!Number.isFinite(size) || size <= 0) {
    return true;
  }
  return size <= getMaxUploadBytes();
}

function assertUploadableFileSize(filename, sizeBytes) {
  if (!isUploadableFileSize(sizeBytes)) {
    throw buildFileTooLargeError(filename, sizeBytes);
  }
}

function isRetryableUploadError(error) {
  const status = error?.response?.status;
  if ([408, 429, 500, 502, 503, 504].includes(status)) return true;

  const code = error?.code;
  return ['ECONNRESET', 'ETIMEDOUT', 'ECONNABORTED', 'ENOTFOUND', 'EAI_AGAIN'].includes(code);
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function uploadToSharePoint({ projectId, projectName, filename, contentType, buffer }) {
  const uploadUrl = powerAutomate.uploadUrl();
  const sharePointPath = buildSharePointPath(projectName, filename);
  if (!buffer) {
    throw new Error(`Missing file buffer for upload: ${filename}`);
  }

  assertUploadableFileSize(filename, buffer.length);
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
    timeoutMs: getUploadTimeoutMs(),
  });

  try {
    let response;
    let lastError;
    const timeout = getUploadTimeoutMs();

    for (let attempt = 1; attempt <= MAX_RETRIES; attempt += 1) {
      try {
        response = await axios.post(uploadUrl, payload, {
          headers: { 'Content-Type': 'application/json' },
          timeout,
          validateStatus: () => true,
        });

        if (response) {
          break;
        }
      } catch (attemptError) {
        lastError = attemptError;
        if (attempt < MAX_RETRIES && isRetryableUploadError(attemptError)) {
          const backoffMs = 1000 * attempt;
          log('Retrying Power Automate upload after error', {
            projectId,
            filename,
            attempt,
            maxAttempts: MAX_RETRIES,
            code: attemptError.code,
            message: attemptError.message,
            backoffMs,
          });
          await delay(backoffMs);
          continue;
        }
        throw attemptError;
      }
    }

    if (!response && lastError) {
      throw lastError;
    }
    if (!response) {
      throw new Error('Power Automate upload failed: no response received');
    }

    const body = response.data || {};
    if (response.status !== 200) {
      log('Treating upload as success because HTTP response was received', {
        projectId,
        filename,
        status: response.status,
      });
    }

    console.log('Upload successful');
    log('Upload successful', {
      projectId,
      filename,
      sharePointPath,
      message: body.message || `File upload accepted (status ${response.status})`,
    });

    return {
      success: true,
      message: body.message || `File upload accepted (status ${response.status})`,
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
  const filename = path.basename(filePath);
  const stats = fs.statSync(filePath);
  assertUploadableFileSize(filename, stats.size);

  const buffer = fs.readFileSync(filePath);

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
  getMaxUploadBytes,
  isUploadableFileSize,
  formatMegabytes,
};
