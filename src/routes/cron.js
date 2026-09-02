const express = require('express');
const scheduleService = require('../services/schedule.service');
const { validateEnv, validateSharePointEnv } = require('../config/env');
const { log, logError } = require('../utils/logger');
const { isVercel } = require('../config/runtime');

const router = express.Router();

function authorizeCron(req, res, next) {
  const secret = process.env.CRON_SECRET;
  if (!secret) {
    if (isVercel()) {
      return res.status(401).json({ success: false, error: 'CRON_SECRET is not configured' });
    }
    return next();
  }

  const header = String(req.headers.authorization || '');
  if (header !== `Bearer ${secret}`) {
    return res.status(401).json({ success: false, error: 'Unauthorized cron request' });
  }
  return next();
}

function enqueueContinue(req) {
  const secret = process.env.CRON_SECRET;
  const host = process.env.VERCEL_URL;
  if (!secret || !host) {
    return;
  }

  const depth = Number(req.headers['x-sync-chain'] || 0);
  if (!Number.isFinite(depth) || depth >= 250) {
    log('Scheduled sync chain depth limit reached', { depth });
    return;
  }

  const url = `https://${host}/api/cron/sync`;
  const promise = fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${secret}`,
      'x-sync-chain': String(depth + 1),
    },
  }).then((response) => {
    if (!response.ok) {
      logError('Failed to continue scheduled sync chain', new Error(`status ${response.status}`));
    }
  }).catch((error) => logError('Failed to continue scheduled sync chain', error));

  try {
    const { waitUntil } = require('@vercel/functions');
    waitUntil(promise);
  } catch {
    promise.catch(() => {});
  }
}

async function handleCron(req, res) {
  try {
    validateEnv();
    validateSharePointEnv();

    const chained = Number(req.headers['x-sync-chain'] || 0) > 0;
    const result = await scheduleService.runScheduledSync({
      continueRun: chained,
    });

    if (result?.incomplete) {
      enqueueContinue(req);
    }

    res.json({
      success: true,
      ...result,
    });
  } catch (error) {
    logError('Cron sync failed', error);
    res.status(500).json({
      success: false,
      error: error.message,
    });
  }
}

router.get('/sync', authorizeCron, handleCron);
router.post('/sync', authorizeCron, handleCron);

module.exports = router;
