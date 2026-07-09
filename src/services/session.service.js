const crypto = require('crypto');

/** @type {{ token: string, username: string, createdAt: number, lastSeen: number } | null} */
let activeSession = null;

function generateToken() {
  return crypto.randomBytes(32).toString('hex');
}

function getActiveSessionInfo() {
  if (!activeSession) return null;
  return { username: activeSession.username };
}

function hasActiveSession() {
  return activeSession !== null;
}

function validateSession(token) {
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
  activeSession = {
    token: generateToken(),
    username,
    createdAt: Date.now(),
    lastSeen: Date.now(),
  };
  return activeSession;
}

function destroySession(token) {
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
};
