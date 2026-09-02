const fs = require('fs');
const path = require('path');
const { isServerless } = require('../config/runtime');
const { dataDir, ensureDir } = require('../config/paths');
const persistentJson = require('./persistentJson.service');

const SETTINGS_FILE = path.join(dataDir(), 'settings.json');
const SETTINGS_RELATIVE = 'data/settings.json';

const EDITABLE_KEYS = [
  'FILEVINE_CLIENT_ID',
  'FILEVINE_CLIENT_SECRET',
  'FILEVINE_PAT',
  'FILEVINE_ORG_ID',
  'FILEVINE_USER_ID',
  'AZURE_TENANT_ID',
  'AZURE_CLIENT_ID',
  'AZURE_CLIENT_SECRET',
  'SHAREPOINT_SITE_ID',
  'SHAREPOINT_DRIVE_ID',
  'SHAREPOINT_ROOT_FOLDER',
];

const READONLY_KEYS = ['FILEVINE_TOKEN_URL', 'FILEVINE_API'];

const DEFAULTS = {
  FILEVINE_TOKEN_URL: 'https://identity.filevine.com/connect/token',
  FILEVINE_API: 'https://api.filevineapp.com/fv-app/v2',
  SHAREPOINT_ROOT_FOLDER: 'Filevine',
};

const API_FIELD_MAP = {
  filevineClientId: 'FILEVINE_CLIENT_ID',
  filevineClientSecret: 'FILEVINE_CLIENT_SECRET',
  filevinePat: 'FILEVINE_PAT',
  filevineOrgId: 'FILEVINE_ORG_ID',
  filevineUserId: 'FILEVINE_USER_ID',
  azureTenantId: 'AZURE_TENANT_ID',
  azureClientId: 'AZURE_CLIENT_ID',
  azureClientSecret: 'AZURE_CLIENT_SECRET',
  sharepointSiteId: 'SHAREPOINT_SITE_ID',
  sharepointDriveId: 'SHAREPOINT_DRIVE_ID',
  sharepointRootFolder: 'SHAREPOINT_ROOT_FOLDER',
  filevineTokenUrl: 'FILEVINE_TOKEN_URL',
  filevineApi: 'FILEVINE_API',
};

/** @type {Record<string, string>} */
let store = {};
let ready = false;
let readyPromise = null;

function loadFromDisk() {
  if (isServerless()) {
    return;
  }

  if (!fs.existsSync(SETTINGS_FILE)) {
    store = {};
    return;
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf8'));
    store = parsed && typeof parsed === 'object' ? parsed : {};
  } catch {
    store = {};
  }
}

async function ensureReady() {
  if (ready) {
    return;
  }
  if (readyPromise) {
    return readyPromise;
  }

  readyPromise = (async () => {
    if (isServerless()) {
      const remote = await persistentJson.read(SETTINGS_RELATIVE);
      if (remote && typeof remote === 'object') {
        store = remote;
      }
    } else {
      loadFromDisk();
    }
    ready = true;
  })().finally(() => {
    readyPromise = null;
  });

  return readyPromise;
}

function get(key) {
  const saved = store[key];
  if (saved !== undefined && saved !== '') {
    return saved;
  }

  const fromEnv = process.env[key];
  if (fromEnv !== undefined && fromEnv !== '') {
    return fromEnv;
  }

  if (DEFAULTS[key] !== undefined) {
    return DEFAULTS[key];
  }

  return undefined;
}

function getPublicSettings() {
  return {
    filevineClientId: get('FILEVINE_CLIENT_ID') || '',
    filevineClientSecret: get('FILEVINE_CLIENT_SECRET') || '',
    filevinePat: get('FILEVINE_PAT') || '',
    filevineOrgId: get('FILEVINE_ORG_ID') || '',
    filevineUserId: get('FILEVINE_USER_ID') || '',
    azureTenantId: get('AZURE_TENANT_ID') || '',
    azureClientId: get('AZURE_CLIENT_ID') || '',
    azureClientSecret: get('AZURE_CLIENT_SECRET') || '',
    sharepointSiteId: get('SHAREPOINT_SITE_ID') || '',
    sharepointDriveId: get('SHAREPOINT_DRIVE_ID') || '',
    sharepointRootFolder: get('SHAREPOINT_ROOT_FOLDER') || DEFAULTS.SHAREPOINT_ROOT_FOLDER,
    filevineTokenUrl: get('FILEVINE_TOKEN_URL') || DEFAULTS.FILEVINE_TOKEN_URL,
    filevineApi: get('FILEVINE_API') || DEFAULTS.FILEVINE_API,
  };
}

async function updateFromApiPayload(payload = {}) {
  const nextStore = { ...store };

  for (const [apiField, envKey] of Object.entries(API_FIELD_MAP)) {
    if (!EDITABLE_KEYS.includes(envKey)) {
      continue;
    }
    if (!Object.prototype.hasOwnProperty.call(payload, apiField)) {
      continue;
    }

    const value = String(payload[apiField] ?? '').trim();
    if (!value) {
      delete nextStore[envKey];
      continue;
    }
    nextStore[envKey] = value;
  }

  store = nextStore;

  if (isServerless()) {
    await persistentJson.write(SETTINGS_RELATIVE, store);
  } else {
    ensureDir(dataDir());
    fs.writeFileSync(SETTINGS_FILE, JSON.stringify(store, null, 2), 'utf8');
  }

  return getPublicSettings();
}

if (!isServerless()) {
  loadFromDisk();
  ready = true;
}

module.exports = {
  EDITABLE_KEYS,
  READONLY_KEYS,
  get,
  getPublicSettings,
  updateFromApiPayload,
  ensureReady,
};
