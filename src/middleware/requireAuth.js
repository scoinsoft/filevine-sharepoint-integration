const sessionService = require('../services/session.service');

function extractToken(req) {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith('Bearer ')) {
    return authHeader.slice(7).trim();
  }
  const headerToken = req.headers['x-session-token'];
  return typeof headerToken === 'string' ? headerToken.trim() : null;
}

function requireAuth(req, res, next) {
  const token = extractToken(req);
  const session = sessionService.validateSession(token);
  if (!session) {
    return res.status(401).json({
      success: false,
      error: 'Not authenticated. Please log in.',
    });
  }
  req.session = session;
  next();
}

module.exports = {
  requireAuth,
  extractToken,
};
