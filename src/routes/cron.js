const express = require('express');

const router = express.Router();

function handleCron(_req, res) {
  res.json({
    success: true,
    skipped: true,
    reason: 'disabled',
    incomplete: false,
    message: 'Scheduled cron sync is disabled.',
  });
}

router.get('/sync', handleCron);
router.post('/sync', handleCron);

module.exports = router;
