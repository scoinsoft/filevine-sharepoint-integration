const path = require('path');
const express = require('express');
const morgan = require('morgan');
const { port, validateEnv } = require('./config/env');
const syncRoutes = require('./routes/sync');

validateEnv();

const app = express();

app.use(morgan('dev'));
app.use(express.json({ limit: '2mb' }));
app.use(express.static(path.join(__dirname, '..', 'public')));

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.use('/api', syncRoutes);

app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    success: false,
    error: err.message,
    details: err.stack,
  });
});

app.listen(port, () => {
  console.log(`Server listening on http://localhost:${port}`);
  console.log(`UI:            http://localhost:${port}/`);
  console.log(`Projects API:  GET  http://localhost:${port}/api/projects`);
  console.log(`Sync stream:   POST http://localhost:${port}/api/projects/:id/sync`);
  console.log(`Legacy test:   GET  http://localhost:${port}/api/test-sync`);
});
