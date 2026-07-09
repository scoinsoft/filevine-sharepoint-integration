const express = require('express');
const settingsService = require('../services/settings.service');
const filevineService = require('../services/filevine.service');

const router = express.Router();

router.get('/', (req, res) => {
  res.json({
    success: true,
    settings: settingsService.getPublicSettings(),
  });
});

router.put('/', (req, res) => {
  try {
    const settings = settingsService.updateFromApiPayload(req.body || {});
    filevineService.clearCachedToken();

    res.json({
      success: true,
      settings,
      message: 'Settings saved.',
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message || 'Failed to save settings',
    });
  }
});

module.exports = router;
