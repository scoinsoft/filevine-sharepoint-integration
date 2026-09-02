const crypto = require('crypto');
const { isServerless } = require('../config/runtime');

const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;

/** In-process session for local single-user lock. Unused on serverless. */
let activeSession = null;

function getSessionSecret() {
  if (process.env.SESSION_SECRET) {
    return process.env.SESSION_SECRET;
  }
  return `fv-sp:${process.env.APP_USERNAME || ''}:${process.env.APP_PASSWORD || ''}`;
}

function signPayload(payload) {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', getSessionSecret()).update(body).digest('base64url');
  return `${body}.${sig}`;
}

function parseToken(token) {
  if (!token || typeof token !== 'string' || !token.includes('.')) {
    return null;
  }
  const [body, sig] = token.split('.');
  if (!body || !sig) {
    return null;
  }

  const expected = crypto.createHmac('sha256', getSessionSecret()).update(body).digest('base64url');
  const sigBuffer = Buffer.from(sig);
  const expectedBuffer = Buffer.from(expected);
  if (sigBuffer.length !== expectedBuffer.length || !crypto.timingSafeEqual(sigBuffer, expectedBuffer)) {
    return null;
  }

  try {
    const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    if (!payload?.u || !payload.exp || Date.now() > payload.exp) {
      return null;
    }
    return payload;
  } catch {
    return null;
  }
}

function generateToken() {
  return crypto.randomBytes(32).toString('hex');
}

function getActiveSessionInfo() {
  if (isServerless()) {
    return null;
  }
  if (!activeSession) return null;
  return { username: activeSession.username };
}

function hasActiveSession() {
  if (isServerless()) {
    return false;
  }
  return activeSession !== null;
}

function validateSession(token) {
  const signed = parseToken(token);
  if (signed) {
    return {
      token,
      username: signed.u,
    };
  }

  if (isServerless()) {
    return null;
  }

  if (!token || !activeSession || activeSession.token !== token) {
    return null;
  }
  activeSession.lastSeen = Date.now();
  return {
    token: activeSession.token,
    username: activeSession.username,
  };
}

function createSession(username) {
  const token = signPayload({
    u: username,
    iat: Date.now(),
    exp: Date.now() + SESSION_TTL_MS,
  });

  if (!isServerless()) {
    activeSession = {
      token,
      username,
      createdAt: Date.now(),
      lastSeen: Date.now(),
    };
    return activeSession;
  }

  return { token, username };
}

function destroySession(token) {
  if (isServerless()) {
    return Boolean(parseToken(token));
  }
  if (!activeSession) return false;
  if (token && activeSession.token !== token) return false;
  activeSession = null;
  return true;
}

function forceDestroySession() {
  activeSession = null;
}

module.exports = {
  getActiveSessionInfo,
  hasActiveSession,
  validateSession,
  createSession,
  destroySession,
  forceDestroySession,
  generateToken,
};
