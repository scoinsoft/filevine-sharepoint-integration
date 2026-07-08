require('dotenv').config();

const filevineRequired = [
  'FILEVINE_CLIENT_ID',
  'FILEVINE_CLIENT_SECRET',
  'FILEVINE_PAT',
  'FILEVINE_ORG_ID',
  'FILEVINE_USER_ID',
];

function getEnv(name, fallback) {
  const value = process.env[name] ?? fallback;
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

function validatePowerAutomateEnv() {
  getEnv('POWER_AUTOMATE_UPLOAD_URL');
}

function validateEnv() {
  validateFilevineEnv();
}

module.exports = {
  port: Number(process.env.PORT) || 3000,
  filevine: {
    clientId: () => getEnv('FILEVINE_CLIENT_ID'),
    clientSecret: () => getEnv('FILEVINE_CLIENT_SECRET'),
    pat: () => getEnv('FILEVINE_PAT'),
    orgId: () => getEnv('FILEVINE_ORG_ID'),
    userId: () => getEnv('FILEVINE_USER_ID'),
    tokenUrl: () => process.env.FILEVINE_TOKEN_URL || 'https://identity.filevine.com/connect/token',
    apiBase: () => process.env.FILEVINE_API || 'https://api.filevineapp.com/fv-app/v2',
  },
  powerAutomate: {
    uploadUrl: () => getEnv('POWER_AUTOMATE_UPLOAD_URL'),
  },
  validateEnv,
  validatePowerAutomateEnv,
};
