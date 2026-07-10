require('dotenv').config();

const settingsService = require('../services/settings.service');

const filevineRequired = [
  'FILEVINE_CLIENT_ID',
  'FILEVINE_CLIENT_SECRET',
  'FILEVINE_PAT',
  'FILEVINE_ORG_ID',
  'FILEVINE_USER_ID',
];

const sharepointRequired = [
  'AZURE_TENANT_ID',
  'AZURE_CLIENT_ID',
  'AZURE_CLIENT_SECRET',
  'SHAREPOINT_SITE_ID',
  'SHAREPOINT_DRIVE_ID',
];

function getEnv(name, fallback) {
  const value = settingsService.get(name) ?? fallback;
  if (value === undefined || value === '') {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function validateFilevineEnv() {
  for (const key of filevineRequired) {
    getEnv(key);
  }
}

function validateSharePointEnv() {
  for (const key of sharepointRequired) {
    getEnv(key);
  }
}

function validateAuthEnv() {
  const username = process.env.APP_USERNAME;
  const password = process.env.APP_PASSWORD;
  if (!username || !password) {
    throw new Error('Missing required environment variable: APP_USERNAME or APP_PASSWORD');
  }
}

function validateEnv() {
  validateFilevineEnv();
  validateAuthEnv();
}

module.exports = {
  port: Number(process.env.PORT) || 3000,
  auth: {
    username: () => process.env.APP_USERNAME,
    password: () => process.env.APP_PASSWORD,
  },
  filevine: {
    clientId: () => getEnv('FILEVINE_CLIENT_ID'),
    clientSecret: () => getEnv('FILEVINE_CLIENT_SECRET'),
    pat: () => getEnv('FILEVINE_PAT'),
    orgId: () => getEnv('FILEVINE_ORG_ID'),
    userId: () => getEnv('FILEVINE_USER_ID'),
    tokenUrl: () => settingsService.get('FILEVINE_TOKEN_URL') || 'https://identity.filevine.com/connect/token',
    apiBase: () => settingsService.get('FILEVINE_API') || 'https://api.filevineapp.com/fv-app/v2',
  },
  sharepoint: {
    tenantId: () => getEnv('AZURE_TENANT_ID'),
    clientId: () => getEnv('AZURE_CLIENT_ID'),
    clientSecret: () => getEnv('AZURE_CLIENT_SECRET'),
    siteId: () => getEnv('SHAREPOINT_SITE_ID'),
    driveId: () => getEnv('SHAREPOINT_DRIVE_ID'),
    rootFolder: () => settingsService.get('SHAREPOINT_ROOT_FOLDER') || process.env.SHAREPOINT_ROOT_FOLDER || 'Filevine',
  },
  validateEnv,
  validateSharePointEnv,
};
