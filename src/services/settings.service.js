const fs = require('fs');
const path = require('path');

const SETTINGS_DIR = path.join(process.cwd(), 'data');
const SETTINGS_FILE = path.join(SETTINGS_DIR, 'settings.json');

const EDITABLE_KEYS = [
  'FILEVINE_CLIENT_ID',
  'FILEVINE_CLIENT_SECRET',
  'FILEVINE_PAT',
  'FILEVINE_ORG_ID',
  'FILEVINE_USER_ID',
  'POWER_AUTOMATE_UPLOAD_URL',
];

const READONLY_KEYS = ['FILEVINE_TOKEN_URL', 'FILEVINE_API'];

const DEFAULTS = {
  FILEVINE_TOKEN_URL: 'https://identity.filevine.com/connect/token',
  FILEVINE_API: 'https://api.filevineapp.com/fv-app/v2',
};

const API_FIELD_MAP = {
  filevineClientId: 'FILEVINE_CLIENT_ID',
  filevineClientSecret: 'FILEVINE_CLIENT_SECRET',
  filevinePat: 'FILEVINE_PAT',
  filevineOrgId: 'FILEVINE_ORG_ID',
  filevineUserId: 'FILEVINE_USER_ID',
  powerAutomateUploadUrl: 'POWER_AUTOMATE_UPLOAD_URL',
  filevineTokenUrl: 'FILEVINE_TOKEN_URL',
  filevineApi: 'FILEVINE_API',
};

/** @type {Record<string, string>} */
let store = {};

function ensureDir() {
  if (!fs.existsSync(SETTINGS_DIR)) {
    fs.mkdirSync(SETTINGS_DIR, { recursive: true });
  }
}

function load() {
  ensureDir();
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
    powerAutomateUploadUrl: get('POWER_AUTOMATE_UPLOAD_URL') || '',
    filevineTokenUrl: get('FILEVINE_TOKEN_URL') || DEFAULTS.FILEVINE_TOKEN_URL,
    filevineApi: get('FILEVINE_API') || DEFAULTS.FILEVINE_API,
  };
}

function updateFromApiPayload(payload = {}) {
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
  ensureDir();
  fs.writeFileSync(SETTINGS_FILE, JSON.stringify(store, null, 2), 'utf8');
  return getPublicSettings();
}

load();

module.exports = {
  EDITABLE_KEYS,
  READONLY_KEYS,
  get,
  getPublicSettings,
  updateFromApiPayload,
};
