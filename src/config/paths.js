const fs = require('fs');
const os = require('os');
const path = require('path');
const { isServerless } = require('./runtime');

function getDataRoot() {
  if (isServerless()) {
    return path.join(os.tmpdir(), 'filevine-sharepoint-sync');
  }
  return process.cwd();
}

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
  return dirPath;
}

function downloadsDir() {
  return path.join(getDataRoot(), 'downloads');
}

function uploadHistoryDir() {
  return path.join(getDataRoot(), 'upload_history');
}

function failedHistoryDir() {
  return path.join(getDataRoot(), 'failed_history');
}

function syncHistoryDir() {
  return path.join(getDataRoot(), 'sync_history');
}

function dataDir() {
  return path.join(getDataRoot(), 'data');
}

module.exports = {
  getDataRoot,
  ensureDir,
  downloadsDir,
  uploadHistoryDir,
  failedHistoryDir,
  syncHistoryDir,
  dataDir,
};
