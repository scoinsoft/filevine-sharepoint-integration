const fs = require('fs');
const path = require('path');
const { isServerless } = require('../config/runtime');
const { getDataRoot, ensureDir } = require('../config/paths');
const { logError } = require('../utils/logger');

function localPath(relativePath) {
  return path.join(getDataRoot(), relativePath);
}

function getSharePoint() {
  return require('./sharepoint.service');
}

async function read(relativePath) {
  const filePath = localPath(relativePath);

  if (fs.existsSync(filePath)) {
    try {
      const parsed = JSON.parse(fs.readFileSync(filePath, 'utf8'));
      if (parsed && typeof parsed === 'object') {
        return parsed;
      }
    } catch (error) {
      logError(`Failed to parse local JSON ${relativePath}`, error);
    }
  }

  if (!isServerless()) {
    return null;
  }

  try {
    return await getSharePoint().readStateJson(relativePath);
  } catch (error) {
    logError(`Failed to read SharePoint state ${relativePath}`, error);
    return null;
  }
}

async function write(relativePath, data) {
  const filePath = localPath(relativePath);
  ensureDir(path.dirname(filePath));
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');

  if (!isServerless()) {
    return filePath;
  }

  try {
    await getSharePoint().writeStateJson(relativePath, data);
  } catch (error) {
    logError(`Failed to persist SharePoint state ${relativePath}`, error);
  }

  return filePath;
}

async function remove(relativePath) {
  const filePath = localPath(relativePath);
  try {
    if (fs.existsSync(filePath)) {
      fs.unlinkSync(filePath);
    }
  } catch (error) {
    logError(`Failed to delete local JSON ${relativePath}`, error);
  }

  if (!isServerless()) {
    return;
  }

  try {
    await getSharePoint().deleteStateJson(relativePath);
  } catch (error) {
    logError(`Failed to delete SharePoint state ${relativePath}`, error);
  }
}

module.exports = {
  read,
  write,
  remove,
  localPath,
};
