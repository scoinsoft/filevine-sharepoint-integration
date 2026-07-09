const fs = require('fs');
const path = require('path');
const cron = require('node-cron');
const syncProjectService = require('./syncProject.service');
const syncHistoryService = require('./syncHistory.service');
const { log, logError } = require('../utils/logger');

const SCHEDULE_DIR = path.join(process.cwd(), 'data');
const SCHEDULE_FILE = path.join(SCHEDULE_DIR, 'schedule.json');
const RUN_WINDOW_MS = 3 * 60 * 60 * 1000;

const DAY_LABELS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

const TIMEZONE_OPTIONS = [
  { value: 'America/New_York', label: 'Eastern Time (US & Canada)' },
  { value: 'America/Chicago', label: 'Central Time (US & Canada)' },
  { value: 'America/Denver', label: 'Mountain Time (US & Canada)' },
  { value: 'America/Los_Angeles', label: 'Pacific Time (US & Canada)' },
  { value: 'America/Phoenix', label: 'Arizona' },
  { value: 'America/Anchorage', label: 'Alaska' },
  { value: 'Pacific/Honolulu', label: 'Hawaii' },
  { value: 'America/Puerto_Rico', label: 'Atlantic Time (Puerto Rico)' },
  { value: 'UTC', label: 'UTC' },
  { value: 'Europe/London', label: 'London' },
  { value: 'Europe/Paris', label: 'Central Europe' },
  { value: 'Asia/Dubai', label: 'Dubai' },
  { value: 'Asia/Kolkata', label: 'India' },
  { value: 'Asia/Singapore', label: 'Singapore' },
  { value: 'Asia/Tokyo', label: 'Tokyo' },
  { value: 'Australia/Sydney', label: 'Sydney' },
];

function getSystemTimezone() {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || process.env.TZ || 'UTC';
  } catch {
    return process.env.TZ || 'UTC';
  }
}

function isValidTimezone(timezone) {
  try {
    Intl.DateTimeFormat(undefined, { timeZone: timezone });
    return true;
  } catch {
    return false;
  }
}

function getTimezoneGmtOffset(timezone) {
  try {
    const formatter = new Intl.DateTimeFormat('en-US', {
      timeZone: timezone,
      timeZoneName: 'shortOffset',
    });
    const offsetPart = formatter
      .formatToParts(new Date())
      .find((part) => part.type === 'timeZoneName');

    const value = offsetPart?.value || 'GMT';
    if (value === 'GMT' || value === 'UTC') {
      return 'GMT+0';
    }
    return value.replace(/^UTC/, 'GMT');
  } catch {
    return 'GMT+0';
  }
}

function enrichTimezoneOption(option) {
  const gmtOffset = getTimezoneGmtOffset(option.value);
  const baseLabel = option.label || option.value;
  return {
    ...option,
    gmtOffset,
    label: `${baseLabel} (${gmtOffset})`,
  };
}

function getTimezoneOptions() {
  const systemTimezone = getSystemTimezone();
  const options = TIMEZONE_OPTIONS.map(enrichTimezoneOption);

  if (!options.some((option) => option.value === systemTimezone)) {
    options.unshift(
      enrichTimezoneOption({
        value: systemTimezone,
        label: `${systemTimezone} (server system)`,
      })
    );
  }

  return {
    systemTimezone: enrichTimezoneOption({
      value: systemTimezone,
      label: `${systemTimezone} (server system)`,
    }),
    options,
  };
}

function resolveTimezone(inputTimezone) {
  const { systemTimezone } = getTimezoneOptions();
  const candidate = typeof inputTimezone === 'string' ? inputTimezone.trim() : '';
  if (candidate && isValidTimezone(candidate)) {
    return candidate;
  }
  return systemTimezone;
}

/** @type {import('node-cron').ScheduledTask | null} */
let cronTask = null;

/** @type {{ active: boolean, startedAt: string | null, estimatedEndAt: string | null, currentProjectName: string | null }} */
let runState = {
  active: false,
  startedAt: null,
  estimatedEndAt: null,
  currentProjectName: null,
};

/** @type {{ enabled: boolean, frequency: string, time: string, dayOfWeek: number, timezone: string }} */
let schedule = getDefaultSchedule();

function getDefaultSchedule() {
  return {
    enabled: false,
    frequency: 'daily',
    time: '02:00',
    dayOfWeek: 1,
    timezone: getSystemTimezone(),
  };
}

function ensureDir() {
  if (!fs.existsSync(SCHEDULE_DIR)) {
    fs.mkdirSync(SCHEDULE_DIR, { recursive: true });
  }
}

function normalizeSchedule(input = {}) {
  const next = getDefaultSchedule();
  if (typeof input.enabled === 'boolean') next.enabled = input.enabled;

  if (input.frequency === 'daily' || input.frequency === 'weekly') {
    next.frequency = input.frequency;
  }

  if (typeof input.time === 'string' && /^\d{2}:\d{2}$/.test(input.time)) {
    next.time = input.time;
  }

  const day = Number(input.dayOfWeek);
  if (Number.isInteger(day) && day >= 0 && day <= 6) {
    next.dayOfWeek = day;
  }

  if (typeof input.timezone === 'string' && input.timezone.trim()) {
    next.timezone = resolveTimezone(input.timezone);
  } else {
    next.timezone = getSystemTimezone();
  }

  return next;
}

function load() {
  ensureDir();
  if (!fs.existsSync(SCHEDULE_FILE)) {
    schedule = getDefaultSchedule();
    return;
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(SCHEDULE_FILE, 'utf8'));
    schedule = normalizeSchedule(parsed);
  } catch {
    schedule = getDefaultSchedule();
  }
}

function save(nextSchedule) {
  schedule = normalizeSchedule(nextSchedule);
  ensureDir();
  fs.writeFileSync(SCHEDULE_FILE, JSON.stringify(schedule, null, 2), 'utf8');
  restartCron();
  return getPublicSchedule();
}

function buildCronExpression(currentSchedule) {
  const [hour, minute] = currentSchedule.time.split(':').map((part) => Number(part));
  if (currentSchedule.frequency === 'weekly') {
    return `${minute} ${hour} * * ${currentSchedule.dayOfWeek}`;
  }
  return `${minute} ${hour} * * *`;
}

function formatScheduleSummary(currentSchedule) {
  const [hour, minute] = currentSchedule.time.split(':');
  const hour12 = Number(hour) % 12 || 12;
  const ampm = Number(hour) >= 12 ? 'PM' : 'AM';
  const timeLabel = `${hour12}:${minute} ${ampm}`;

  if (currentSchedule.frequency === 'weekly') {
    return `Weekly on ${DAY_LABELS[currentSchedule.dayOfWeek]} at ${timeLabel}`;
  }
  return `Daily at ${timeLabel}`;
}

function isUploadBlocked() {
  return runState.active;
}

function getRunStatus() {
  return {
    active: runState.active,
    startedAt: runState.startedAt,
    estimatedEndAt: runState.estimatedEndAt,
    currentProjectName: runState.currentProjectName,
    uploadsBlocked: runState.active,
  };
}

function getTimezoneLabel(timezone) {
  const { options } = getTimezoneOptions();
  const match = options.find((option) => option.value === timezone);
  return match?.label || timezone;
}

function getPublicSchedule() {
  return {
    ...schedule,
    summary: formatScheduleSummary(schedule),
    dayLabel: DAY_LABELS[schedule.dayOfWeek],
    timezoneLabel: getTimezoneLabel(schedule.timezone),
  };
}

async function runScheduledSync() {
  if (runState.active) {
    log('Scheduled sync skipped because a run is already active');
    return;
  }

  const runStartedAt = new Date().toISOString();
  const scheduledRunId = runStartedAt;
  const runFolder = syncHistoryService.createSyncRunFolder(runStartedAt, 'scheduled');

  runState.active = true;
  runState.startedAt = runStartedAt;
  runState.estimatedEndAt = new Date(Date.now() + RUN_WINDOW_MS).toISOString();
  runState.currentProjectName = null;

  log('Scheduled sync started', {
    startedAt: runState.startedAt,
    estimatedEndAt: runState.estimatedEndAt,
    scheduledRunId,
    runFolder: runFolder.relativeDir,
  });

  const projectEntries = [];

  try {
    const projects = await syncProjectService.listAllProjects();
    log(`Scheduled sync processing ${projects.length} project(s)`);

    for (const project of projects) {
      runState.currentProjectName = project.projectName;

      try {
        const summary = await syncProjectService.syncProject(project.projectId, project.projectName, {
          trigger: 'scheduled',
          scheduledRunId,
          runFolder,
        });

        projectEntries.push({
          projectId: project.projectId,
          projectName: project.projectName,
          record: {
            success: summary.success,
            error: summary.error || null,
            counts: summary.counts,
          },
          historyFile: summary.historyFile || null,
        });
      } catch (error) {
        logError(`Scheduled sync failed for project ${project.projectId}`, error);
        const failedSummary = error.syncSummary || null;
        projectEntries.push({
          projectId: project.projectId,
          projectName: project.projectName,
          error: error.message,
          record: failedSummary
            ? {
                success: false,
                error: failedSummary.error || error.message,
                counts: failedSummary.counts,
              }
            : null,
          historyFile: error.historyFile || failedSummary?.historyFile || null,
        });
      }
    }

    const runFinishedAt = new Date().toISOString();
    syncHistoryService.saveScheduledRunHistory(
      {
        runId: scheduledRunId,
        runFolder,
        startedAt: runStartedAt,
        finishedAt: runFinishedAt,
        durationMs: Math.max(0, new Date(runFinishedAt).getTime() - new Date(runStartedAt).getTime()),
      },
      projectEntries
    );

    log('Scheduled sync completed', { projects: projects.length });
  } catch (error) {
    logError('Scheduled sync failed', error);

    const runFinishedAt = new Date().toISOString();
    syncHistoryService.saveScheduledRunHistory(
      {
        runId: scheduledRunId,
        runFolder,
        startedAt: runStartedAt,
        finishedAt: runFinishedAt,
        durationMs: Math.max(0, new Date(runFinishedAt).getTime() - new Date(runStartedAt).getTime()),
      },
      projectEntries
    );
  } finally {
    runState.active = false;
    runState.currentProjectName = null;
    runState.estimatedEndAt = null;
  }
}

function stopCron() {
  if (cronTask) {
    cronTask.stop();
    cronTask = null;
  }
}

function restartCron() {
  stopCron();
  if (!schedule.enabled) return;

  const expression = buildCronExpression(schedule);
  if (!cron.validate(expression)) {
    logError('Invalid cron expression for schedule', new Error(expression));
    return;
  }

  cronTask = cron.schedule(expression, () => {
    runScheduledSync().catch((error) => logError('Scheduled sync task failed', error));
  }, {
    timezone: schedule.timezone,
  });

  log('Schedule cron started', {
    expression,
    timezone: schedule.timezone,
    summary: formatScheduleSummary(schedule),
  });
}

function init() {
  load();
  restartCron();
}

init();

module.exports = {
  DAY_LABELS,
  getPublicSchedule,
  getRunStatus,
  getSystemTimezone,
  getTimezoneOptions,
  isUploadBlocked,
  save,
  runScheduledSync,
  restartCron,
};
