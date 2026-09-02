const fs = require('fs');
const path = require('path');
const persistentJson = require('./persistentJson.service');
const { isServerless } = require('../config/runtime');
const { dataDir, ensureDir } = require('../config/paths');
const { log } = require('../utils/logger');

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
    timezone: 'America/Denver',
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
    next.timezone = 'America/Denver';
  }

  return next;
}

function applyPersisted(parsed) {
  schedule = normalizeSchedule({ ...parsed, enabled: false });
  runState = emptyRunState();
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
  schedule = normalizeSchedule({ ...nextSchedule, enabled: false });
  await persist();
  stopCron();
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

function shouldStartScheduledRun() {
  return false;
}

async function runScheduledSync() {
  log('Scheduled sync is disabled');
  return { skipped: true, reason: 'disabled', incomplete: false };
}

function stopCron() {
  if (cronTask) {
    cronTask.stop();
    cronTask = null;
  }
}

function restartCron() {
  stopCron();
}

function init() {
  if (isServerless()) {
    return;
  }
  loadFromDisk();
  schedule.enabled = false;
  stopCron();
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
