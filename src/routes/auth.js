const express = require('express');
const { auth } = require('../config/env');
const sessionService = require('../services/session.service');
const { extractToken } = require('../middleware/requireAuth');

const router = express.Router();

router.get('/me', (req, res) => {
  const token = extractToken(req);
  const session = sessionService.validateSession(token);
  if (!session) {
    return res.status(401).json({
      success: false,
      error: 'Not authenticated',
    });
  }
  res.json({
    success: true,
    username: session.username,
  });
});

router.post('/login', (req, res) => {
  const { username, password, force } = req.body || {};

  if (!username || !password) {
    return res.status(400).json({
      success: false,
      error: 'Username and password are required',
    });
  }

  if (username !== auth.username() || password !== auth.password()) {
    return res.status(401).json({
      success: false,
      error: 'Invalid username or password',
    });
  }

  if (sessionService.hasActiveSession() && !force) {
    const active = sessionService.getActiveSessionInfo();
    return res.status(409).json({
      success: false,
      conflict: true,
      error: 'Another user is already logged in',
      activeUser: active?.username || 'another user',
    });
  }

  if (force) {
    sessionService.forceDestroySession();
  }

  const session = sessionService.createSession(username);
  res.json({
    success: true,
    token: session.token,
    username: session.username,
  });
});

router.post('/logout', (req, res) => {
  const token = extractToken(req);
  const destroyed = sessionService.destroySession(token);
  if (!destroyed) {
    return res.status(401).json({
      success: false,
      error: 'Not authenticated',
    });
  }
  res.json({ success: true });
});

module.exports = router;
