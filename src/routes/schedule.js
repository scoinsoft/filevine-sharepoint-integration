const express = require('express');
const scheduleService = require('../services/schedule.service');

const router = express.Router();

router.get('/', (req, res) => {
  const { systemTimezone, options: timezoneOptions } = scheduleService.getTimezoneOptions();
  res.json({
    success: true,
    schedule: scheduleService.getPublicSchedule(),
    run: scheduleService.getRunStatus(),
    dayOptions: scheduleService.DAY_LABELS.map((label, value) => ({ value, label })),
    systemTimezone: systemTimezone.value,
    systemTimezoneLabel: systemTimezone.label,
    systemTimezoneGmt: systemTimezone.gmtOffset,
    timezoneOptions,
  });
});

router.put('/', (req, res) => {
  try {
    const schedule = scheduleService.save(req.body || {});
    res.json({
      success: true,
      schedule,
      run: scheduleService.getRunStatus(),
      message: schedule.enabled ? 'Schedule saved and enabled.' : 'Schedule saved (disabled).',
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message || 'Failed to save schedule',
    });
  }
});

module.exports = router;
