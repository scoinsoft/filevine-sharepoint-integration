require('./config/secrets');
const path = require('path');
const express = require('express');
const morgan = require('morgan');
const { port, validateEnv } = require('./config/env');
const {
  isServerless,
  runWithRequestContext,
  startRequestDeadline,
} = require('./config/runtime');
const authRoutes = require('./routes/auth');
const settingsRoutes = require('./routes/settings');
const scheduleRoutes = require('./routes/schedule');
const syncRoutes = require('./routes/sync');
const cronRoutes = require('./routes/cron');
const { requireAuth } = require('./middleware/requireAuth');
const settingsService = require('./services/settings.service');
const scheduleService = require('./services/schedule.service');

const app = express();

app.set('trust proxy', 1);
app.use(morgan(isServerless() ? 'tiny' : 'dev'));
app.use(express.json({ limit: '2mb' }));

app.use((req, res, next) => {
  runWithRequestContext(startRequestDeadline(), () => next());
});

app.use(async (req, res, next) => {
  if (req.path === '/health') {
    return next();
  }
  try {
    await settingsService.ensureReady();
    await scheduleService.ensureReady();
    next();
  } catch (error) {
    next(error);
  }
});

app.use(express.static(path.join(__dirname, '..', 'public')));

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    runtime: isServerless() ? 'serverless' : 'local',
  });
});

app.use('/api/cron', cronRoutes);
app.use('/api/auth', authRoutes);
app.use('/api/settings', requireAuth, settingsRoutes);
app.use('/api/schedule', requireAuth, scheduleRoutes);
app.use('/api', requireAuth, syncRoutes);

app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    success: false,
    error: err.message,
    details: err.stack,
  });
});

if (require.main === module) {
  validateEnv();
  app.listen(port, () => {
    console.log(`Server listening on http://localhost:${port}`);
    console.log(`UI:            http://localhost:${port}/`);
    console.log(`Projects API:  GET  http://localhost:${port}/api/projects`);
    console.log(`Sync stream:   POST http://localhost:${port}/api/projects/:id/sync`);
  });
}

module.exports = app;
