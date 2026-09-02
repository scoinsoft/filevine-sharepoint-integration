const fs = require('fs');
const path = require('path');
const cron = require('node-cron');
const syncProjectService = require('./syncProject.service');
const syncHistoryService = require('./syncHistory.service');
const persistentJson = require('./persistentJson.service');
const { isServerless, shouldStopForDeadline } = require('../config/runtime');
const { dataDir, ensureDir } = require('../config/paths');
const { log, logError } = require('../utils/logger');

const SCHEDULE_FILE = path.join(dataDir(), 'schedule.json');
const SCHEDULE_RELATIVE = 'data/schedule.json';
const RUN_WINDOW_MS = 3 * 60 * 60 * 1000;
const STALE_RUN_MS = 12 * 60 * 1000;

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
  return systemTimezone.value;
}

/** @type {import('node-cron').ScheduledTask | null} */
let cronTask = null;

/** @type {{ active: boolean, startedAt: string | null, estimatedEndAt: string | null, currentProjectName: string | null, heartbeatAt: string | null, runId: string | null, nextIndex: number, lastCompletedRunAt: string | null }} */
let runState = emptyRunState();

/** @type {{ enabled: boolean, frequency: string, time: string, dayOfWeek: number, timezone: string }} */
let schedule = getDefaultSchedule();
let ready = !isServerless();

function emptyRunState() {
  return {
    active: false,
    startedAt: null,
    estimatedEndAt: null,
    currentProjectName: null,
    heartbeatAt: null,
    runId: null,
    nextIndex: 0,
    lastCompletedRunAt: null,
  };
}

function getDefaultSchedule() {
  return {
    enabled: false,
    frequency: 'daily',
    time: '02:00',
    dayOfWeek: 1,
    timezone: getSystemTimezone(),
  };
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

function applyPersisted(parsed) {
  schedule = normalizeSchedule(parsed);
  const persistedRun = parsed?.run && typeof parsed.run === 'object' ? parsed.run : {};
  runState = {
    ...emptyRunState(),
    ...persistedRun,
    lastCompletedRunAt: persistedRun.lastCompletedRunAt || parsed?.lastCompletedRunAt || null,
  };
}

function persistPayload() {
  return {
    ...schedule,
    run: runState,
    lastCompletedRunAt: runState.lastCompletedRunAt,
  };
}

function loadFromDisk() {
  if (!fs.existsSync(SCHEDULE_FILE)) {
    schedule = getDefaultSchedule();
    return;
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(SCHEDULE_FILE, 'utf8'));
    applyPersisted(parsed);
  } catch {
    schedule = getDefaultSchedule();
  }
}

async function persist() {
  const payload = persistPayload();
  if (isServerless()) {
    await persistentJson.write(SCHEDULE_RELATIVE, payload);
    return;
  }
  ensureDir(dataDir());
  fs.writeFileSync(SCHEDULE_FILE, JSON.stringify(payload, null, 2), 'utf8');
}

async function ensureReady() {
  if (ready) {
    return;
  }
  const remote = await persistentJson.read(SCHEDULE_RELATIVE);
  if (remote) {
    applyPersisted(remote);
  } else {
    schedule = getDefaultSchedule();
  }
  ready = true;
}

async function save(nextSchedule) {
  schedule = normalizeSchedule(nextSchedule);
  await persist();
  if (!isServerless()) {
    restartCron();
  }
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

function isRunStale() {
  if (!runState.active) return false;
  const heartbeat = runState.heartbeatAt || runState.startedAt;
  if (!heartbeat) return true;
  return Date.now() - new Date(heartbeat).getTime() > STALE_RUN_MS;
}

function isUploadBlocked() {
  if (!runState.active) return false;
  if (isRunStale()) return false;
  return true;
}

function getRunStatus() {
  const active = isUploadBlocked();
  return {
    active,
    startedAt: active ? runState.startedAt : null,
    estimatedEndAt: active ? runState.estimatedEndAt : null,
    currentProjectName: active ? runState.currentProjectName : null,
    uploadsBlocked: active,
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

function getZonedParts(date, timezone) {
  const formatter = new Intl.DateTimeFormat('en-US', {
    timeZone: timezone,
    weekday: 'short',
    hour: '2-digit',
    minute: '2-digit',
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  });
  const parts = formatter.formatToParts(date);
  const get = (type) => parts.find((part) => part.type === type)?.value;
  const weekday = get('weekday');
  const dayMap = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };
  return {
    year: get('year'),
    month: get('month'),
    day: get('day'),
    hour: get('hour'),
    minute: get('minute'),
    dayOfWeek: dayMap[weekday] ?? 0,
    dateKey: `${get('year')}-${get('month')}-${get('day')}`,
  };
}

function shouldStartScheduledRun(now = new Date()) {
  if (!schedule.enabled) return false;
  const parts = getZonedParts(now, schedule.timezone || 'UTC');
  if (runState.lastCompletedRunAt) {
    const last = getZonedParts(new Date(runState.lastCompletedRunAt), schedule.timezone || 'UTC');
    if (last.dateKey === parts.dateKey) {
      return false;
    }
  }

  if (schedule.frequency === 'weekly' && parts.dayOfWeek !== schedule.dayOfWeek) {
    return false;
  }

  // Vercel cron invokes this endpoint; run once per local schedule-day when enabled.
  if (isServerless()) {
    return true;
  }

  const [hour, minute] = schedule.time.split(':');
  if (parts.hour !== hour) return false;
  const scheduledMinute = Number(minute);
  const currentMinute = Number(parts.minute);
  if (Number.isNaN(scheduledMinute) || currentMinute < scheduledMinute || currentMinute > scheduledMinute + 9) {
    return false;
  }
  return true;
}

async function runScheduledSync(options = {}) {
  await ensureReady();

  const continueRun = Boolean(options.continueRun);
  const force = Boolean(options.force);

  if (runState.active && !isRunStale() && !continueRun) {
    log('Scheduled sync skipped because a run is already active');
    return { skipped: true, reason: 'already-active', incomplete: false };
  }

  const resumeExisting = Boolean(runState.active && runState.runId && (continueRun || isRunStale()));
  const startNew = force || shouldStartScheduledRun();

  if (!resumeExisting && !startNew) {
    return { skipped: true, reason: 'not-due', incomplete: false };
  }

  const runStartedAt = resumeExisting ? runState.startedAt : new Date().toISOString();
  const scheduledRunId = resumeExisting ? runState.runId : runStartedAt;
  const startIndex = resumeExisting ? Number(runState.nextIndex) || 0 : 0;
  const runFolder = syncHistoryService.createSyncRunFolder(runStartedAt, 'scheduled');

  runState.active = true;
  runState.startedAt = runStartedAt;
  runState.estimatedEndAt = new Date(Date.now() + RUN_WINDOW_MS).toISOString();
  runState.currentProjectName = null;
  runState.heartbeatAt = new Date().toISOString();
  runState.runId = scheduledRunId;
  runState.nextIndex = startIndex;
  await persist();

  log('Scheduled sync started', {
    startedAt: runState.startedAt,
    estimatedEndAt: runState.estimatedEndAt,
    scheduledRunId,
    startIndex,
    runFolder: runFolder.relativeDir,
  });

  const projectEntries = [];
  let incomplete = false;

  try {
    const projects = await syncProjectService.listAllProjects();
    log(`Scheduled sync processing ${projects.length} project(s) from index ${startIndex}`);

    for (let index = startIndex; index < projects.length; index += 1) {
      if (shouldStopForDeadline()) {
        runState.nextIndex = index;
        runState.heartbeatAt = new Date().toISOString();
        incomplete = true;
        await persist();
        log('Scheduled sync pausing for function time budget', { nextIndex: index });
        break;
      }

      const project = projects[index];
      runState.currentProjectName = project.projectName;
      runState.nextIndex = index;
      runState.heartbeatAt = new Date().toISOString();
      await persist();

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
            skippedArchivedProject: Boolean(summary.skippedArchivedProject),
            incomplete: Boolean(summary.incomplete),
            error: summary.error || null,
            counts: summary.counts,
          },
          historyFile: summary.historyFile || null,
        });

        if (summary.incomplete) {
          runState.nextIndex = index;
          incomplete = true;
          await persist();
          break;
        }

        runState.nextIndex = index + 1;
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
        runState.nextIndex = index + 1;
      }
    }

    if (!incomplete && runState.nextIndex >= projects.length) {
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

      runState = {
        ...emptyRunState(),
        lastCompletedRunAt: runFinishedAt,
      };
      await persist();
      log('Scheduled sync completed', { projects: projects.length });
      return { skipped: false, incomplete: false, processed: projectEntries.length };
    }

    return {
      skipped: false,
      incomplete: true,
      nextIndex: runState.nextIndex,
      processed: projectEntries.length,
    };
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
    runState.active = false;
    runState.currentProjectName = null;
    await persist();
    throw error;
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
  if (isServerless() || !schedule.enabled) return;

  const expression = buildCronExpression(schedule);
  if (!cron.validate(expression)) {
    logError('Invalid cron expression for schedule', new Error(expression));
    return;
  }

  cronTask = cron.schedule(expression, () => {
    runScheduledSync({ force: true }).catch((error) => logError('Scheduled sync task failed', error));
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
  if (isServerless()) {
    return;
  }
  loadFromDisk();
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
  ensureReady,
  shouldStartScheduledRun,
};
