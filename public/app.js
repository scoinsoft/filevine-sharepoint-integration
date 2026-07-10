(() => {
  const SESSION_KEY = 'fv_sp_session_token';
  const SYNC_ALL_SNAPSHOT_KEY = 'fv_sp_sync_all_snapshot_v1';
  const SYNC_ALL_PROJECT_CONCURRENCY = 3;
  const MAX_SINGLE_RESULT_ROWS = 600;
  const MAX_SYNC_ALL_HISTORY_ROWS = 200;
  const TRANSFER_TOAST_MIN_INTERVAL_MS = 350;
  const SYNC_ALL_UI_MIN_INTERVAL_MS = 250;

  const state = {
    projects: [],
    selected: null,
    documents: [],
    syncing: false,
    syncingAll: false,
    uploadedProjectIds: new Set(),
    uploadHistoryLoaded: false,
    paging: {
      offset: 0,
      pageSize: 100,
      fetchSize: 1000,
      loadedTo: 0,
      hasMore: false,
      loading: false,
      pageCache: new Map(),
      inFlight: new Map(),
      prefetchRunning: false,
      allLoaded: false,
    },
    syncAll: {
      projectsTotal: 0,
      projectsDone: 0,
      filesOk: 0,
      filesSkipped: 0,
      filesFail: 0,
      projectsSkippedArchived: 0,
      elapsedMs: 0,
      timerStartedAt: null,
      timerIntervalId: null,
      currentProjectName: '',
      activeProjects: [],
      paused: false,
      pausedDueToCrash: false,
      pauseReason: '',
      pauseResolver: null,
      remainingProjects: [],
      processing: false,
      resumeSnapshot: null,
    },
    singleSync: {
      pausedDueToCrash: false,
      pauseReason: '',
    },
    schedule: {
      config: null,
      run: { uploadsBlocked: false, active: false },
      dayOptions: [],
      systemTimezone: 'UTC',
      systemTimezoneLabel: 'UTC (GMT+0)',
      timezoneOptions: [],
      pollTimer: null,
    },
    backgroundSyncs: [],
  };


  const $ = (id) => document.getElementById(id);

  const els = {
    loginView: $('login-view'),
    workView: $('work-view'),
    loginForm: $('login-form'),
    loginUsername: $('login-username'),
    loginPassword: $('login-password'),
    loginError: $('login-error'),
    loginBtn: $('login-btn'),
    loginBtnLabel: $('login-btn-label'),
    loginBtnSpinner: $('login-btn-spinner'),
    loginBtnIcon: $('login-btn-icon'),
    sessionConflictModal: $('session-conflict-modal'),
    sessionConflictText: $('session-conflict-text'),
    sessionConflictCancel: $('session-conflict-cancel'),
    sessionConflictConfirm: $('session-conflict-confirm'),
    loggedInUser: $('logged-in-user'),
    loggedInUserName: $('logged-in-user-name'),
    scheduleBtn: $('schedule-btn'),
    scheduleStatusBar: $('schedule-status-bar'),
    scheduleStatusSummary: $('schedule-status-summary'),
    scheduleStatusMeta: $('schedule-status-meta'),
    scheduleNotice: $('schedule-notice'),
    backgroundSyncNotice: $('background-sync-notice'),
    settingsBtn: $('settings-btn'),
    logoutBtn: $('logout-btn'),
    scheduleModal: $('schedule-modal'),
    scheduleForm: $('schedule-form'),
    scheduleError: $('schedule-error'),
    scheduleSuccess: $('schedule-success'),
    scheduleEnabled: $('schedule-enabled'),
    scheduleFrequencyDaily: $('schedule-frequency-daily'),
    scheduleFrequencyWeekly: $('schedule-frequency-weekly'),
    scheduleDayWrap: $('schedule-day-wrap'),
    scheduleDay: $('schedule-day'),
    scheduleTime: $('schedule-time'),
    scheduleTimezone: $('schedule-timezone'),
    scheduleTimezoneHint: $('schedule-timezone-hint'),
    scheduleSummary: $('schedule-summary'),
    scheduleCancelBtn: $('schedule-cancel-btn'),
    scheduleSaveBtn: $('schedule-save-btn'),
    scheduleSaveSpinner: $('schedule-save-spinner'),
    scheduleSaveIcon: $('schedule-save-icon'),
    scheduleSaveLabel: $('schedule-save-label'),
    settingsModal: $('settings-modal'),
    settingsForm: $('settings-form'),
    settingsError: $('settings-error'),
    settingsSuccess: $('settings-success'),
    settingsCancelBtn: $('settings-cancel-btn'),
    settingsSaveBtn: $('settings-save-btn'),
    settingsSaveSpinner: $('settings-save-spinner'),
    settingsSaveIcon: $('settings-save-icon'),
    settingsSaveLabel: $('settings-save-label'),
    settingsFilevineClientId: $('settings-filevine-client-id'),
    settingsFilevineClientSecret: $('settings-filevine-client-secret'),
    settingsFilevinePat: $('settings-filevine-pat'),
    settingsFilevineOrgId: $('settings-filevine-org-id'),
    settingsFilevineUserId: $('settings-filevine-user-id'),
    settingsAzureTenantId: $('settings-azure-tenant-id'),
    settingsAzureClientId: $('settings-azure-client-id'),
    settingsAzureClientSecret: $('settings-azure-client-secret'),
    settingsSharepointSiteId: $('settings-sharepoint-site-id'),
    settingsSharepointDriveId: $('settings-sharepoint-drive-id'),
    settingsSharepointRootFolder: $('settings-sharepoint-root-folder'),
    settingsFilevineTokenUrl: $('settings-filevine-token-url'),
    settingsFilevineApi: $('settings-filevine-api'),
    error: $('error'),
    pauseNotice: $('pause-notice'),
    projectsStatusSpinner: $('projects-status-spinner'),
    projectsStatusText: $('projects-status-text'),
    refreshProjectsBtn: $('refresh-projects-btn'),
    refreshProjectsIcon: $('refresh-projects-icon'),
    projectsList: $('projects-list'),
    prevProjectsBtn: $('prev-projects-btn'),
    nextProjectsBtn: $('next-projects-btn'),
    projectsPageInfo: $('projects-page-info'),
    filesStatusSpinner: $('files-status-spinner'),
    filesStatusText: $('files-status-text'),
    refreshFilesBtn: $('refresh-files-btn'),
    refreshFilesIcon: $('refresh-files-icon'),
    filesList: $('files-list'),
    syncBtn: $('sync-btn'),
    syncBtnLabel: $('sync-btn-label'),
    syncBtnSpinner: $('sync-btn-spinner'),
    syncBtnIcon: $('sync-btn-icon'),
    progressBox: $('progress-box'),
    progressText: $('progress-text'),
    progressPercent: $('progress-percent'),
    progressBar: $('progress-bar'),
    transferToastPanel: $('transfer-toast-panel'),
    transferToastHeader: $('transfer-toast-header'),
    transferToastStack: $('transfer-toast-stack'),
    transferToastDownloadCount: $('transfer-toast-download-count'),
    transferToastUploadCount: $('transfer-toast-upload-count'),
    statTotal: $('stat-total'),
    statOk: $('stat-ok'),
    statFail: $('stat-fail'),
    resultList: $('result-list'),
    syncAllBtn: $('sync-all-btn'),
    syncAllBtnLabel: $('sync-all-btn-label'),
    syncAllBtnSpinner: $('sync-all-btn-spinner'),
    syncAllBtnIcon: $('sync-all-btn-icon'),
    syncNewBtn: $('sync-new-btn'),
    syncNewBtnLabel: $('sync-new-btn-label'),
    syncNewBtnSpinner: $('sync-new-btn-spinner'),
    syncNewBtnIcon: $('sync-new-btn-icon'),
    pauseSyncAllBtn: $('pause-sync-all-btn'),
    resumeSyncAllBtn: $('resume-sync-all-btn'),
    resumeLastSyncAllBtn: $('resume-last-sync-all-btn'),
    syncAllHintSpinner: $('sync-all-hint-spinner'),
    syncAllHintText: $('sync-all-hint-text'),
    syncAllProgressBox: $('sync-all-progress-box'),
    syncAllProgressText: $('sync-all-progress-text'),
    syncAllProgressPercent: $('sync-all-progress-percent'),
    syncAllElapsed: $('sync-all-elapsed'),
    syncAllProgressBar: $('sync-all-progress-bar'),
    syncAllProjectsDone: $('sync-all-projects-done'),
    syncAllProjectsTotal: $('sync-all-projects-total'),
    syncAllFilesOk: $('sync-all-files-ok'),
    syncAllFilesSkipped: $('sync-all-files-skipped'),
    syncAllProjectsSkippedArchived: $('sync-all-projects-skipped-archived'),
    syncAllFilesFail: $('sync-all-files-fail'),
    syncAllResultList: $('sync-all-result-list'),
  };

  function formatArchivedProjectSkipLine(project, summary = {}) {
    const name = summary.projectName || project?.projectName || 'Unknown project';
    return `Archived: ${name}`;
  }

  function isLikelyArchivedProject(project) {
    return project?.isArchived === true || project?.phaseName === 'Archived';
  }

  function isProjectAlreadyUploaded(project) {
    if (!project) return false;
    if (typeof project.alreadyUploaded === 'boolean') return project.alreadyUploaded;
    return state.uploadedProjectIds.has(String(project.projectId));
  }

  function isNewProject(project) {
    if (!project) return false;
    if (isLikelyArchivedProject(project)) return false;
    if (typeof project.isNew === 'boolean') return project.isNew && !isLikelyArchivedProject(project);
    return !isProjectAlreadyUploaded(project);
  }

  function markProjectUploadedLocally(projectId) {
    if (projectId == null) return;
    const key = String(projectId);
    state.uploadedProjectIds.add(key);
    for (const page of state.paging.pageCache.values()) {
      for (const project of page?.projects || []) {
        if (String(project.projectId) === key) {
          project.alreadyUploaded = true;
          project.isNew = false;
        }
      }
    }
    for (const project of state.projects) {
      if (String(project.projectId) === key) {
        project.alreadyUploaded = true;
        project.isNew = false;
      }
    }
    if (state.selected && String(state.selected.projectId) === key) {
      state.selected.alreadyUploaded = true;
      state.selected.isNew = false;
    }
  }

  function getNewProjects() {
    return getAllCachedProjects().filter((project) => isNewProject(project));
  }

  function getProjectArchiveCounts() {
    const projects = getAllCachedProjects();
    let archived = 0;
    let newCount = 0;
    for (const project of projects) {
      if (isLikelyArchivedProject(project)) {
        archived += 1;
        continue;
      }
      if (!isProjectAlreadyUploaded(project)) {
        newCount += 1;
      }
    }
    return {
      total: projects.length,
      archived,
      active: projects.length - archived,
      newCount,
    };
  }

  function formatProjectTotalsLabel() {
    const { total, archived, active, newCount } = getProjectArchiveCounts();
    if (total === 0) return '';

    const breakdown = ` · ${active} active · ${archived} archived · ${newCount} new`;
    if (state.paging.allLoaded) {
      return ` · ${state.paging.loadedTo} total${breakdown}`;
    }
    return ` · ${total} loaded${breakdown}`;
  }

  function renderProjectStatusRibbon(project) {
    if (!isLikelyArchivedProject(project)) return '';
    return `
      <span class="pointer-events-none absolute right-0 top-0 h-12 w-12 overflow-hidden" aria-hidden="true">
        <span class="absolute right-[-18px] top-[8px] w-[72px] rotate-45 bg-amber-500 py-0.5 text-center text-[9px] font-bold uppercase tracking-wide text-white shadow-sm">
          Archived
        </span>
      </span>
      <span class="sr-only">Archived project</span>
    `;
  }

  function renderProjectNewBadge(project) {
    if (!isNewProject(project)) return '';
    return `<span class="project-new-badge shrink-0">New</span>`;
  }

  function getSessionToken() {
    return localStorage.getItem(SESSION_KEY);
  }

  function setSessionToken(token) {
    localStorage.setItem(SESSION_KEY, token);
  }

  function clearSessionToken() {
    localStorage.removeItem(SESSION_KEY);
  }

  function isIncompleteSyncSnapshot(snapshot) {
    if (!snapshot) return false;
    const total = Number(snapshot.projectsTotal) || 0;
    const done = Number(snapshot.projectsDone) || 0;
    if (total > 0 && done >= total) return false;

    const remaining = Array.isArray(snapshot.remainingProjects) ? snapshot.remainingProjects : [];
    const remainingIds = Array.isArray(snapshot.remainingProjectIds) ? snapshot.remainingProjectIds : [];
    return remaining.length > 0 || remainingIds.length > 0 || (total > 0 && done < total);
  }

  function normalizeSyncAllSnapshot(parsed) {
    if (!parsed || typeof parsed !== 'object') return null;

    let remainingProjects = Array.isArray(parsed.remainingProjects) ? parsed.remainingProjects : [];
    if (!remainingProjects.length && Array.isArray(parsed.remainingProjectIds)) {
      remainingProjects = parsed.remainingProjectIds.map((projectId) => ({
        projectId,
        projectName: `Project ${projectId}`,
      }));
    }

    remainingProjects = remainingProjects
      .map((project) => ({
        projectId: project?.projectId,
        projectName: project?.projectName || `Project ${project?.projectId}`,
        projectNumber: project?.projectNumber ?? null,
        phaseName: project?.phaseName ?? null,
        createdDate: project?.createdDate ?? null,
      }))
      .filter((project) => project.projectId != null);

    const snapshot = {
      ...parsed,
      remainingProjects,
      remainingProjectIds: remainingProjects.map((project) => project.projectId),
    };

    return isIncompleteSyncSnapshot(snapshot) ? snapshot : null;
  }

  let lastSnapshotSaveAt = 0;
  let snapshotSaveTimer = null;

  function saveSyncAllSnapshot(options = {}) {
    if (!state.syncingAll) return;

    const force = Boolean(options.force);
    const now = Date.now();
    if (!force && now - lastSnapshotSaveAt < 5000) {
      if (!snapshotSaveTimer) {
        snapshotSaveTimer = window.setTimeout(() => {
          snapshotSaveTimer = null;
          saveSyncAllSnapshot({ force: true });
        }, 5000);
      }
      return;
    }
    lastSnapshotSaveAt = now;
    if (snapshotSaveTimer) {
      window.clearTimeout(snapshotSaveTimer);
      snapshotSaveTimer = null;
    }

    const remaining = state.syncAll.remainingProjects || [];
    const total = Number(state.syncAll.projectsTotal) || 0;
    const done = Number(state.syncAll.projectsDone) || 0;
    if (!remaining.length && total > 0 && done >= total) return;

    const progressPercent = total > 0 ? Math.round((done / total) * 100) : 0;
    const currentProject =
      state.syncAll.activeProjects?.[0] ||
      state.syncAll.currentProjectName ||
      remaining[0]?.projectName ||
      '';
    // Keep snapshot light: IDs for resume + a small preview of names for the UI.
    const remainingPreview = remaining.slice(0, 25).map((project) => ({
      projectId: project.projectId,
      projectName: project.projectName,
    }));
    const payload = {
      savedAt: new Date().toISOString(),
      projectsTotal: state.syncAll.projectsTotal,
      projectsDone: state.syncAll.projectsDone,
      filesOk: state.syncAll.filesOk,
      filesSkipped: state.syncAll.filesSkipped,
      filesFail: state.syncAll.filesFail,
      projectsSkippedArchived: state.syncAll.projectsSkippedArchived,
      elapsedMs: state.syncAll.elapsedMs,
      progressPercent,
      currentProject,
      remainingProjectIds: remaining.map((project) => project.projectId),
      remainingProjects: remainingPreview,
    };
    try {
      localStorage.setItem(SYNC_ALL_SNAPSHOT_KEY, JSON.stringify(payload));
      state.syncAll.resumeSnapshot = payload;
    } catch (error) {
      console.warn('Could not save sync snapshot (storage full or blocked)', error);
    }
  }

  function clearSyncAllSnapshot() {
    localStorage.removeItem(SYNC_ALL_SNAPSHOT_KEY);
    state.syncAll.resumeSnapshot = null;
  }

  function loadSyncAllSnapshot() {
    const raw = localStorage.getItem(SYNC_ALL_SNAPSHOT_KEY);
    if (!raw) return null;
    try {
      return normalizeSyncAllSnapshot(JSON.parse(raw));
    } catch {
      return null;
    }
  }

  function refreshResumeSnapshot() {
    const snapshot = loadSyncAllSnapshot();
    state.syncAll.resumeSnapshot = snapshot;
    if (!snapshot) {
      localStorage.removeItem(SYNC_ALL_SNAPSHOT_KEY);
    }
    return snapshot;
  }

  function getResumeSnapshot() {
    if (state.syncAll.resumeSnapshot && isIncompleteSyncSnapshot(state.syncAll.resumeSnapshot)) {
      return state.syncAll.resumeSnapshot;
    }
    return refreshResumeSnapshot();
  }

  function resolveSnapshotRemainingProjects(snapshot) {
    const cached = getAllCachedProjects();
    const cacheById = new Map(cached.map((project) => [String(project.projectId), project]));
    const previewById = new Map(
      (snapshot.remainingProjects || []).map((project) => [String(project.projectId), project])
    );

    const remainingIds = snapshot.remainingProjectIds || [];
    if (remainingIds.length > 0) {
      return remainingIds.map((projectId) => {
        const key = String(projectId);
        const cachedProject = cacheById.get(key);
        if (cachedProject) return cachedProject;
        const preview = previewById.get(key);
        return {
          projectId,
          projectName: preview?.projectName || `Project ${projectId}`,
        };
      });
    }

    const fromSnapshot = (snapshot.remainingProjects || []).map((project) => {
      const cachedProject = cacheById.get(String(project.projectId));
      return (
        cachedProject || {
          projectId: project.projectId,
          projectName: project.projectName || `Project ${project.projectId}`,
        }
      );
    });

    if (fromSnapshot.length > 0) {
      return fromSnapshot;
    }

    const total = Number(snapshot.projectsTotal) || 0;
    const done = Number(snapshot.projectsDone) || 0;
    if (cached.length >= total && total > done) {
      return cached.slice(done);
    }

    return [];
  }

  function formatElapsedMs(ms) {
    const totalSeconds = Math.max(0, Math.floor((Number(ms) || 0) / 1000));
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return [hours, minutes, seconds].map((v) => String(v).padStart(2, '0')).join(':');
  }

  function updateSyncAllElapsedLabel() {
    if (!els.syncAllElapsed) return;
    els.syncAllElapsed.textContent = formatElapsedMs(state.syncAll.elapsedMs);
  }

  function stopSyncAllTimer() {
    if (state.syncAll.timerIntervalId) {
      clearInterval(state.syncAll.timerIntervalId);
      state.syncAll.timerIntervalId = null;
    }
    state.syncAll.timerStartedAt = null;
  }

  function startSyncAllTimer() {
    if (!state.syncingAll || state.syncAll.paused || state.syncAll.timerIntervalId) return;

    state.syncAll.timerStartedAt = Date.now();
    state.syncAll.timerIntervalId = setInterval(() => {
      if (!state.syncingAll || state.syncAll.paused) return;
      const now = Date.now();
      const delta = Math.max(0, now - (state.syncAll.timerStartedAt || now));
      state.syncAll.timerStartedAt = now;
      state.syncAll.elapsedMs += delta;
      updateSyncAllElapsedLabel();
    }, 1000);
  }

  function authHeaders(extra = {}) {
    const headers = { ...extra };
    const token = getSessionToken();
    if (token) {
      headers.Authorization = `Bearer ${token}`;
    }
    return headers;
  }

  function isUploadBlockedBySchedule() {
    return Boolean(state.schedule.run?.uploadsBlocked);
  }

  function updateScheduleNotice() {
    if (!isUploadBlockedBySchedule()) {
      els.scheduleNotice.classList.add('hidden');
      els.scheduleNotice.textContent = '';
      return;
    }

    const projectNote = state.schedule.run.currentProjectName
      ? ` Currently syncing: ${state.schedule.run.currentProjectName}.`
      : '';
    els.scheduleNotice.textContent =
      `Scheduled sync is running and usually takes about 2–3 hours.${projectNote} Manual uploads are disabled until it finishes.`;
    els.scheduleNotice.classList.remove('hidden');
  }

  function isProjectSyncingOnServer(projectId) {
    return state.backgroundSyncs.some(
      (run) => String(run.projectId) === String(projectId)
    );
  }

  function updateBackgroundSyncNotice() {
    if (!state.backgroundSyncs.length) {
      els.backgroundSyncNotice.classList.add('hidden');
      els.backgroundSyncNotice.textContent = '';
      return;
    }

    const names = state.backgroundSyncs.map((run) => run.projectName).join(', ');
    els.backgroundSyncNotice.textContent =
      `Upload still running on the server for: ${names}. Refreshing the browser does not stop it — already-uploaded files are saved and will be skipped on retry.`;
    els.backgroundSyncNotice.classList.remove('hidden');
  }

  async function refreshSyncStatus() {
    try {
      const response = await apiFetch('/api/sync/status');
      const data = await response.json().catch(() => ({}));
      if (response.ok && data.success) {
        state.backgroundSyncs = data.activeRuns || [];
        updateBackgroundSyncNotice();
        updateSyncButton();
        updateSyncAllControls();
      }
    } catch {
      // Ignore polling errors.
    }
  }

  function applyScheduleStatus(data = {}) {
    state.schedule.config = data.schedule || null;
    state.schedule.run = data.run || { uploadsBlocked: false, active: false };
    if (Array.isArray(data.dayOptions)) {
      state.schedule.dayOptions = data.dayOptions;
    }
    if (typeof data.systemTimezone === 'string' && data.systemTimezone) {
      state.schedule.systemTimezone = data.systemTimezone;
    }
    if (typeof data.systemTimezoneLabel === 'string' && data.systemTimezoneLabel) {
      state.schedule.systemTimezoneLabel = data.systemTimezoneLabel;
    }
    if (Array.isArray(data.timezoneOptions)) {
      state.schedule.timezoneOptions = data.timezoneOptions;
    }
    updateScheduleButtonRibbon();
    updateScheduleNotice();
    updateSyncButton();
    updateSyncAllControls();
  }

  function getScheduleStatusDisplay(schedule = {}) {
    if (!schedule?.enabled) {
      return null;
    }

    const summary = schedule.summary || 'Scheduled';
    const gmtMatch = schedule.timezoneLabel?.match(/\(GMT[^)]+\)/);
    const gmt = gmtMatch ? gmtMatch[0] : '';
    const timezoneName =
      schedule.timezoneLabel?.replace(/\s*\(GMT[^)]+\)\s*$/, '').trim() ||
      schedule.timezone ||
      '';

    return {
      summary,
      meta: [gmt, timezoneName].filter(Boolean).join(' · '),
      title: schedule.timezoneLabel ? `${summary} — ${schedule.timezoneLabel}` : summary,
    };
  }

  function updateScheduleButtonRibbon() {
    const display = getScheduleStatusDisplay(state.schedule.config);

    if (!display) {
      els.scheduleStatusBar.classList.add('hidden');
      els.scheduleStatusSummary.textContent = '';
      els.scheduleStatusMeta.textContent = '';
      els.scheduleStatusBar.title = '';
      els.scheduleBtn.classList.remove('nav-btn-active');
      return;
    }

    els.scheduleStatusSummary.textContent = display.summary;
    els.scheduleStatusMeta.textContent = display.meta;
    els.scheduleStatusBar.title = display.title;
    els.scheduleStatusBar.classList.remove('hidden');
    els.scheduleBtn.classList.add('nav-btn-active');
  }

  function getDefaultScheduleTimezone() {
    return state.schedule.systemTimezone || 'UTC';
  }

  function fillTimezoneOptions(selectedTimezone) {
    const timezone = selectedTimezone || getDefaultScheduleTimezone();
    const options = state.schedule.timezoneOptions.length
      ? [...state.schedule.timezoneOptions]
      : [{ value: getDefaultScheduleTimezone(), label: `${getDefaultScheduleTimezone()} (server system)` }];

    if (timezone && !options.some((option) => option.value === timezone)) {
      options.push({ value: timezone, label: timezone });
    }

    els.scheduleTimezone.innerHTML = options
      .map(
        (option) =>
          `<option value="${escapeHtml(option.value)}"${
            option.value === timezone ? ' selected' : ''
          }>${escapeHtml(option.label)}</option>`
      )
      .join('');

    els.scheduleTimezoneHint.textContent = `Server system timezone: ${state.schedule.systemTimezoneLabel || getDefaultScheduleTimezone()}. The schedule runs at the selected time in the timezone you choose.`;
  }

  async function refreshScheduleStatus() {
    try {
      const response = await apiFetch('/api/schedule');
      const data = await response.json().catch(() => ({}));
      if (response.ok && data.success) {
        applyScheduleStatus(data);
      }
    } catch {
      // Ignore polling errors; auth handler will redirect if needed.
    }
  }

  function startSchedulePolling() {
    stopSchedulePolling();
    refreshScheduleStatus();
    refreshSyncStatus();
    state.schedule.pollTimer = window.setInterval(() => {
      refreshScheduleStatus();
      refreshSyncStatus();
    }, 20000);
  }

  function stopSchedulePolling() {
    if (state.schedule.pollTimer) {
      window.clearInterval(state.schedule.pollTimer);
      state.schedule.pollTimer = null;
    }
  }

  function updateScheduleFrequencyUi() {
    const weekly = els.scheduleFrequencyWeekly.checked;
    els.scheduleDayWrap.classList.toggle('hidden', !weekly);
  }

  function fillScheduleDayOptions(selectedDay = 1) {
    const options = state.schedule.dayOptions.length
      ? state.schedule.dayOptions
      : [
          { value: 0, label: 'Sunday' },
          { value: 1, label: 'Monday' },
          { value: 2, label: 'Tuesday' },
          { value: 3, label: 'Wednesday' },
          { value: 4, label: 'Thursday' },
          { value: 5, label: 'Friday' },
          { value: 6, label: 'Saturday' },
        ];

    els.scheduleDay.innerHTML = options
      .map(
        (option) =>
          `<option value="${option.value}"${Number(option.value) === Number(selectedDay) ? ' selected' : ''}>${option.label}</option>`
      )
      .join('');
  }

  function fillScheduleForm(schedule = {}) {
    els.scheduleEnabled.checked = Boolean(schedule.enabled);
    if (schedule.frequency === 'weekly') {
      els.scheduleFrequencyWeekly.checked = true;
    } else {
      els.scheduleFrequencyDaily.checked = true;
    }
    els.scheduleTime.value = schedule.time || '02:00';
    fillTimezoneOptions(schedule.timezone || getDefaultScheduleTimezone());
    fillScheduleDayOptions(schedule.dayOfWeek ?? 1);
    updateScheduleFrequencyUi();
    const timezoneLabel = schedule.timezoneLabel || schedule.timezone || getDefaultScheduleTimezone();
    els.scheduleSummary.textContent = schedule.enabled
      ? `${schedule.summary || 'Scheduled'} (${timezoneLabel})`
      : 'Not scheduled';
  }

  function readScheduleForm() {
    return {
      enabled: els.scheduleEnabled.checked,
      frequency: els.scheduleFrequencyWeekly.checked ? 'weekly' : 'daily',
      time: els.scheduleTime.value || '02:00',
      dayOfWeek: Number(els.scheduleDay.value),
      timezone: els.scheduleTimezone.value || getDefaultScheduleTimezone(),
    };
  }

  function showScheduleError(message) {
    if (!message) {
      els.scheduleError.classList.add('hidden');
      els.scheduleError.textContent = '';
      return;
    }
    els.scheduleError.textContent = message;
    els.scheduleError.classList.remove('hidden');
  }

  function showScheduleSuccess(message) {
    if (!message) {
      els.scheduleSuccess.classList.add('hidden');
      els.scheduleSuccess.textContent = '';
      return;
    }
    els.scheduleSuccess.textContent = message;
    els.scheduleSuccess.classList.remove('hidden');
  }

  function setScheduleLoading(loading) {
    els.scheduleSaveBtn.disabled = loading;
    els.scheduleCancelBtn.disabled = loading;
    els.scheduleSaveSpinner.classList.toggle('hidden', !loading);
    els.scheduleSaveIcon.classList.toggle('hidden', loading);
    els.scheduleSaveLabel.textContent = loading ? 'Saving…' : 'Save schedule';
  }

  function openScheduleModal() {
    els.scheduleModal.classList.remove('hidden');
    els.scheduleModal.classList.add('flex');
    showScheduleError('');
    showScheduleSuccess('');
    loadSchedule();
  }

  function closeScheduleModal() {
    els.scheduleModal.classList.add('hidden');
    els.scheduleModal.classList.remove('flex');
    showScheduleError('');
    showScheduleSuccess('');
  }

  async function loadSchedule() {
    els.scheduleSaveBtn.disabled = true;
    showScheduleError('');
    showScheduleSuccess('');

    try {
      const response = await apiFetch('/api/schedule');
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to load schedule');
      }
      applyScheduleStatus(data);
      fillScheduleForm(data.schedule || {});
    } catch (error) {
      showScheduleError(error.message || 'Could not load schedule.');
      fillTimezoneOptions(getDefaultScheduleTimezone());
      fillScheduleDayOptions();
      updateScheduleFrequencyUi();
    } finally {
      els.scheduleSaveBtn.disabled = false;
    }
  }

  async function saveSchedule(event) {
    event.preventDefault();
    showScheduleError('');
    showScheduleSuccess('');
    setScheduleLoading(true);

    try {
      const response = await apiFetch('/api/schedule', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(readScheduleForm()),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to save schedule');
      }
      applyScheduleStatus(data);
      fillScheduleForm(data.schedule || {});
      showScheduleSuccess(data.message || 'Schedule saved.');
    } catch (error) {
      showScheduleError(error.message || 'Could not save schedule.');
    } finally {
      setScheduleLoading(false);
    }
  }

  function showLoginPage() {
    stopSchedulePolling();
    els.workView.classList.add('hidden');
    els.loginView.classList.remove('hidden');
    els.loggedInUserName.textContent = '';
    hideLoginError();
    hideSessionConflictModal();
  }

  function showWorkPage(username) {
    els.loginView.classList.add('hidden');
    els.workView.classList.remove('hidden');
    els.loggedInUserName.textContent = username ? username : '';
    hideLoginError();
    hideSessionConflictModal();
    refreshResumeSnapshot();
    updateSyncAllControls();
    startSchedulePolling();
  }

  function showLoginError(message) {
    if (!message) {
      els.loginError.classList.add('hidden');
      els.loginError.textContent = '';
      return;
    }
    els.loginError.textContent = message;
    els.loginError.classList.remove('hidden');
  }

  function hideLoginError() {
    showLoginError('');
  }

  function showSessionConflictModal(activeUser) {
    const label = activeUser || 'another user';
    els.sessionConflictText.textContent =
      `"${label}" is already signed in. Only one user can use this automation tool at a time. Would you like to sign them out and continue?`;
    els.sessionConflictModal.classList.remove('hidden');
    els.sessionConflictModal.classList.add('flex');
  }

  function hideSessionConflictModal() {
    els.sessionConflictModal.classList.add('hidden');
    els.sessionConflictModal.classList.remove('flex');
  }

  function setLoginLoading(loading) {
    els.loginBtn.disabled = loading;
    els.loginUsername.disabled = loading;
    els.loginPassword.disabled = loading;
    els.loginBtnSpinner.classList.toggle('hidden', !loading);
    els.loginBtnIcon.classList.toggle('hidden', loading);
    els.loginBtnLabel.textContent = loading ? 'Signing in…' : 'Sign in';
  }

  function setProjectsStatus(text, loading = false) {
    els.projectsStatusText.textContent = text;
    els.projectsStatusSpinner.classList.toggle('hidden', !loading);
  }

  function setFilesStatus(text, loading = false) {
    els.filesStatusText.textContent = text;
    els.filesStatusSpinner.classList.toggle('hidden', !loading);
  }

  function setSyncAllHint(text, loading = false) {
    els.syncAllHintText.textContent = text;
    els.syncAllHintSpinner.classList.toggle('hidden', !loading);
  }

  async function apiFetch(url, options = {}) {
    const response = await fetch(url, {
      ...options,
      headers: authHeaders(options.headers || {}),
    });

    if (response.status === 401 && getSessionToken()) {
      clearSessionToken();
      showLoginPage();
      throw new Error('Session expired. Please sign in again.');
    }

    return response;
  }

  async function restoreSession() {
    const token = getSessionToken();
    if (!token) {
      showLoginPage();
      return false;
    }

    try {
      const response = await fetch('/api/auth/me', { headers: authHeaders() });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.success) {
        clearSessionToken();
        showLoginPage();
        return false;
      }

      showWorkPage(data.username);
      return true;
    } catch {
      clearSessionToken();
      showLoginPage();
      return false;
    }
  }

  async function handleLogin(force = false) {
    const username = els.loginUsername.value.trim();
    const password = els.loginPassword.value;

    if (!username || !password) {
      showLoginError('Username and password are required.');
      return;
    }

    hideLoginError();
    setLoginLoading(true);

    try {
      const response = await fetch('/api/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password, force }),
      });
      const data = await response.json().catch(() => ({}));

      if (response.status === 409 && data.conflict) {
        showSessionConflictModal(data.activeUser);
        return;
      }

      if (!response.ok || !data.success) {
        showLoginError(data.error || 'Sign in failed.');
        return;
      }

      setSessionToken(data.token);
      els.loginPassword.value = '';
      showWorkPage(data.username);
      await loadUploadHistoryIndex({ rebuild: true });
      loadProjects();
    } catch (error) {
      showLoginError(error.message || 'Could not reach the server.');
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleLogout() {
    try {
      await fetch('/api/auth/logout', {
        method: 'POST',
        headers: authHeaders({ 'Content-Type': 'application/json' }),
      });
    } catch {
      // Still clear local session if server is unreachable.
    }

    clearSessionToken();
    showLoginPage();
    els.loginPassword.value = '';
  }

  function showSettingsError(message) {
    if (!message) {
      els.settingsError.classList.add('hidden');
      els.settingsError.textContent = '';
      return;
    }
    els.settingsError.textContent = message;
    els.settingsError.classList.remove('hidden');
  }

  function showSettingsSuccess(message) {
    if (!message) {
      els.settingsSuccess.classList.add('hidden');
      els.settingsSuccess.textContent = '';
      return;
    }
    els.settingsSuccess.textContent = message;
    els.settingsSuccess.classList.remove('hidden');
  }

  function setSettingsLoading(loading) {
    els.settingsSaveBtn.disabled = loading;
    els.settingsCancelBtn.disabled = loading;
    els.settingsSaveSpinner.classList.toggle('hidden', !loading);
    els.settingsSaveIcon.classList.toggle('hidden', loading);
    els.settingsSaveLabel.textContent = loading ? 'Saving…' : 'Save settings';
  }

  function fillSettingsForm(settings = {}) {
    els.settingsFilevineClientId.value = settings.filevineClientId || '';
    els.settingsFilevineClientSecret.value = settings.filevineClientSecret || '';
    els.settingsFilevinePat.value = settings.filevinePat || '';
    els.settingsFilevineOrgId.value = settings.filevineOrgId || '';
    els.settingsFilevineUserId.value = settings.filevineUserId || '';
    els.settingsAzureTenantId.value = settings.azureTenantId || '';
    els.settingsAzureClientId.value = settings.azureClientId || '';
    els.settingsAzureClientSecret.value = settings.azureClientSecret || '';
    els.settingsSharepointSiteId.value = settings.sharepointSiteId || '';
    els.settingsSharepointDriveId.value = settings.sharepointDriveId || '';
    els.settingsSharepointRootFolder.value = settings.sharepointRootFolder || 'Filevine';
    els.settingsFilevineTokenUrl.value = settings.filevineTokenUrl || '';
    els.settingsFilevineApi.value = settings.filevineApi || '';
  }

  function readSettingsForm() {
    return {
      filevineClientId: els.settingsFilevineClientId.value.trim(),
      filevineClientSecret: els.settingsFilevineClientSecret.value.trim(),
      filevinePat: els.settingsFilevinePat.value.trim(),
      filevineOrgId: els.settingsFilevineOrgId.value.trim(),
      filevineUserId: els.settingsFilevineUserId.value.trim(),
      azureTenantId: els.settingsAzureTenantId.value.trim(),
      azureClientId: els.settingsAzureClientId.value.trim(),
      azureClientSecret: els.settingsAzureClientSecret.value.trim(),
      sharepointSiteId: els.settingsSharepointSiteId.value.trim(),
      sharepointDriveId: els.settingsSharepointDriveId.value.trim(),
      sharepointRootFolder: els.settingsSharepointRootFolder.value.trim() || 'Filevine',
    };
  }

  function openSettingsModal() {
    els.settingsModal.classList.remove('hidden');
    els.settingsModal.classList.add('flex');
    showSettingsError('');
    showSettingsSuccess('');
    loadSettings();
  }

  function closeSettingsModal() {
    els.settingsModal.classList.add('hidden');
    els.settingsModal.classList.remove('flex');
    showSettingsError('');
    showSettingsSuccess('');
  }

  async function loadSettings() {
    showSettingsError('');
    showSettingsSuccess('');
    els.settingsSaveBtn.disabled = true;

    try {
      const response = await apiFetch('/api/settings');
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to load settings');
      }
      fillSettingsForm(data.settings || {});
    } catch (error) {
      showSettingsError(error.message || 'Could not load settings.');
    } finally {
      els.settingsSaveBtn.disabled = false;
    }
  }

  async function saveSettings(event) {
    event.preventDefault();
    showSettingsError('');
    showSettingsSuccess('');
    setSettingsLoading(true);

    try {
      const response = await apiFetch('/api/settings', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(readSettingsForm()),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to save settings');
      }

      fillSettingsForm(data.settings || {});
      showSettingsSuccess(data.message || 'Settings saved.');
    } catch (error) {
      showSettingsError(error.message || 'Could not save settings.');
    } finally {
      setSettingsLoading(false);
    }
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function parseNonNegativeInt(value, fallback = 0) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed) || parsed < 0) {
      return fallback;
    }
    return Math.floor(parsed);
  }

  function showError(message) {
    if (!message) {
      els.error.classList.add('hidden');
      els.error.textContent = '';
      return;
    }
    els.error.textContent = message;
    els.error.classList.remove('hidden');
  }

  function showPauseNotice(message) {
    if (!message) {
      els.pauseNotice.classList.add('hidden');
      els.pauseNotice.textContent = '';
      return;
    }
    els.pauseNotice.textContent = message;
    els.pauseNotice.classList.remove('hidden');
  }

  function hidePauseNotice() {
    showPauseNotice('');
  }

  function isBackendConnectionError(error) {
    const message = String(error?.message || error || '').toLowerCase();
    return (
      error instanceof TypeError ||
      message.includes('failed to fetch') ||
      message.includes('network') ||
      message.includes('connection lost') ||
      message.includes('backend connection') ||
      message.includes('connection reset') ||
      message.includes('econnreset') ||
      message.includes('socket hang up') ||
      message.includes('aborted')
    );
  }

  function getBackendPauseMessage(error) {
    const detail = error?.message ? ` (${error.message})` : '';
    return (
      `Sync paused: backend connection was lost${detail}. ` +
      'Restart the server if it crashed, then click Resume. Already-uploaded files will be skipped.'
    );
  }

  function setButtonLoading(spinnerEl, labelEl, loading, idleText, loadingText, iconEl) {
    if (loading) {
      spinnerEl.classList.remove('hidden');
      if (iconEl) iconEl.classList.add('hidden');
      labelEl.textContent = loadingText;
    } else {
      spinnerEl.classList.add('hidden');
      if (iconEl) iconEl.classList.remove('hidden');
      labelEl.textContent = idleText;
    }
  }

  function getAllCachedProjects() {
    const projects = [];
    const seen = new Set();

    const offsets = [...state.paging.pageCache.keys()].sort((a, b) => a - b);
    for (const offset of offsets) {
      const page = state.paging.pageCache.get(offset);
      for (const project of page?.projects || []) {
        const key = String(project.projectId);
        if (seen.has(key)) continue;
        seen.add(key);
        projects.push(project);
      }
    }

    return projects;
  }

  function updateSyncButton() {
    const scheduledBlocked = isUploadBlockedBySchedule();
    const serverSyncingSelected =
      state.selected && isProjectSyncingOnServer(state.selected.projectId);
    const busy = state.syncing || (state.syncingAll && !state.syncAll.paused);
    const canResumeSingle = state.singleSync.pausedDueToCrash && state.selected;
    els.syncBtn.disabled =
      scheduledBlocked ||
      serverSyncingSelected ||
      ((!state.selected || busy || state.documents.length === 0) && !canResumeSingle);

    if (state.syncing) {
      setButtonLoading(
        els.syncBtnSpinner,
        els.syncBtnLabel,
        true,
        "Send project files to SharePoint",
        'Sending to SharePoint…',
        els.syncBtnIcon
      );
    } else if (canResumeSingle) {
      els.syncBtnSpinner.classList.add('hidden');
      els.syncBtnIcon.classList.remove('hidden');
      els.syncBtnLabel.textContent = 'Resume upload';
    } else if (scheduledBlocked) {
      els.syncBtnSpinner.classList.add('hidden');
      els.syncBtnIcon.classList.remove('hidden');
      els.syncBtnLabel.textContent = 'Send project files to SharePoint';
    } else {
      setButtonLoading(
        els.syncBtnSpinner,
        els.syncBtnLabel,
        false,
        "Send project files to SharePoint",
        'Sending to SharePoint…',
        els.syncBtnIcon
      );
    }
    updateRefreshButtons();
  }

  function updateSyncAllControls() {
    const count = state.paging.loadedTo;
    const ready = state.paging.allLoaded && count > 0;
    const scheduledBlocked = isUploadBlockedBySchedule();
    const serverSyncing = state.backgroundSyncs.length > 0;
    const busy = state.syncing || (state.syncingAll && !state.syncAll.paused);
    const paused = state.syncingAll && state.syncAll.paused;
    const canResumeAll = state.syncingAll && paused;
    const resumeSnapshot = getResumeSnapshot();
    const canResumePrevious = !state.syncingAll && !state.syncing && Boolean(resumeSnapshot);

    els.syncAllBtn.disabled = scheduledBlocked || serverSyncing || !ready || busy;
    setButtonLoading(
      els.syncAllBtnSpinner,
      els.syncAllBtnLabel,
      state.syncingAll && !paused,
      `Sync all to SharePoint (${count})`,
      `Syncing all (${count})…`,
      els.syncAllBtnIcon
    );

    const newCount = ready ? getNewProjects().length : 0;
    if (els.syncNewBtn) {
      els.syncNewBtn.disabled =
        scheduledBlocked || serverSyncing || !ready || busy || newCount === 0;
      setButtonLoading(
        els.syncNewBtnSpinner,
        els.syncNewBtnLabel,
        false,
        newCount > 0 ? `Sync new projects (${newCount})` : 'Sync new projects',
        'Syncing new projects…',
        els.syncNewBtnIcon
      );
      els.syncNewBtn.title =
        newCount > 0
          ? `${newCount} active project(s) not yet in upload history`
          : 'No new projects available';
    }

    els.pauseSyncAllBtn.disabled = !state.syncingAll || paused;
    els.resumeSyncAllBtn.disabled = !canResumeAll;
    els.resumeLastSyncAllBtn.disabled = !canResumePrevious;
    if (canResumePrevious) {
      const snap = resumeSnapshot;
      const remainingCount = snap.remainingProjects?.length || snap.remainingProjectIds?.length || 0;
      const projectLabel = snap.currentProject || snap.remainingProjects?.[0]?.projectName || 'unknown project';
      const tooltip = `Last interrupted at ${snap.progressPercent ?? 0}% (${snap.projectsDone}/${snap.projectsTotal} projects). ${remainingCount} remaining. Next: ${projectLabel}.`;
      els.resumeLastSyncAllBtn.title = tooltip;
      els.resumeLastSyncAllBtn.setAttribute('aria-label', tooltip);
    } else {
      els.resumeLastSyncAllBtn.title = '';
      els.resumeLastSyncAllBtn.setAttribute('aria-label', 'Resume previous run');
    }

    if (state.syncingAll && state.syncAll.pausedDueToCrash) {
      setSyncAllHint('Paused due to backend/network issue. Fix the server, then click Resume.', false);
    } else if (state.syncingAll && paused) {
      setSyncAllHint(
        state.syncAll.currentProjectName
          ? `Paused — now finishing ${state.syncAll.currentProjectName}. Click Resume to continue.`
          : 'Paused. Click Resume to continue.',
        false
      );
    } else if (state.syncingAll) {
      setSyncAllHint(
        state.syncAll.currentProjectName
          ? `Syncing ${state.syncAll.currentProjectName}…`
          : 'Syncing all projects…',
        true
      );
    } else if (serverSyncing) {
      setSyncAllHint('A project upload is still running on the server.', true);
    } else if (canResumePrevious) {
      const remainingCount =
        resumeSnapshot.remainingProjects?.length || resumeSnapshot.remainingProjectIds?.length || 0;
      setSyncAllHint(
        `Previous run interrupted at ${resumeSnapshot.progressPercent ?? 0}% (${resumeSnapshot.projectsDone}/${resumeSnapshot.projectsTotal}). ${remainingCount} project(s) ready to resume.`,
        false
      );
    } else if (scheduledBlocked) {
      setSyncAllHint('Scheduled sync is running — manual uploads are disabled (about 2–3 hours).', true);
    } else if (!state.paging.allLoaded) {
      setSyncAllHint(`Loading projects… (${count} loaded)`, true);
    } else if (count === 0) {
      setSyncAllHint('No projects available to sync.', false);
    } else if (newCount > 0) {
      setSyncAllHint(
        `${newCount} new active project(s) available. Use Sync new projects to upload only those.`,
        false
      );
    } else {
      setSyncAllHint(`Ready — ${count} project(s) loaded.`, false);
    }
  }

  function updateSyncAllButton() {
    updateSyncAllControls();
  }

  async function waitIfPaused() {
    if (!state.syncAll.paused) return;

    addSyncAllHistory('Paused — waiting for Resume…', 'info');
    updateSyncAllControls();
    updateSyncAllProgressUi();

    await new Promise((resolve) => {
      state.syncAll.pauseResolver = resolve;
    });
    state.syncAll.pauseResolver = null;
  }

  function isSharePointConfigError(error) {
    return Boolean(
      error?.sharePointConfigError ||
      String(error?.message || '').includes('SharePoint is misconfigured') ||
      String(error?.message || '').includes('SharePoint authentication failed')
    );
  }

  function getPauseMessage(error) {
    if (isSharePointConfigError(error)) {
      return (
        error?.message ||
        'SharePoint is misconfigured. Fix AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET in .env, then restart the server.'
      );
    }
    return getBackendPauseMessage(error);
  }

  function pauseSyncAllDueToCrash(error, remainingProjects) {
    if (!state.syncingAll) return;

    const message = getPauseMessage(error);
    state.syncAll.paused = true;
    state.syncAll.pausedDueToCrash = true;
    state.syncAll.pauseReason = message;
    state.syncAll.remainingProjects = remainingProjects || [];
    stopSyncAllTimer();
    saveSyncAllSnapshot();
    state.syncAll.currentProjectName = '';

    showPauseNotice(message);
    showError('');
    addSyncAllHistory(`Paused due to backend issue: ${error?.message || 'connection lost'}`, 'error');
    updateSyncAllControls();
    updateSyncAllProgressUi();
  }

  function pauseSyncAll() {
    if (!state.syncingAll || state.syncAll.paused) return;
    state.syncAll.paused = true;
    stopSyncAllTimer();
    saveSyncAllSnapshot();
    addSyncAllHistory(
      'Pause requested — current project will finish, then sync waits.',
      'info'
    );
    updateSyncAllControls();
    updateSyncAllProgressUi();
  }

  function resumeSyncAll() {
    if (!state.syncingAll || !state.syncAll.paused) return;

    state.syncAll.paused = false;
    startSyncAllTimer();
    if (state.syncAll.pausedDueToCrash) {
      state.syncAll.pausedDueToCrash = false;
      state.syncAll.pauseReason = '';
      hidePauseNotice();
      addSyncAllHistory('Resumed sync after backend pause.', 'info');
      updateSyncAllControls();
      updateSyncAllProgressUi();
      continueSyncAll(state.syncAll.remainingProjects);
      return;
    }

    addSyncAllHistory('Resumed sync.', 'info');
    if (state.syncAll.pauseResolver) {
      state.syncAll.pauseResolver();
    }
    updateSyncAllControls();
    updateSyncAllProgressUi();
  }

  function updateProjectsPager() {
    const start = state.projects.length === 0 ? 0 : state.paging.offset + 1;
    const end = state.paging.offset + state.projects.length;
    els.projectsPageInfo.textContent =
      state.projects.length === 0 ? '0-0' : `${start}-${end}`;
    els.prevProjectsBtn.disabled =
      state.paging.loading || state.syncingAll || state.paging.offset === 0;
    els.nextProjectsBtn.disabled =
      state.paging.loading || state.syncingAll || !canGoNextPage();
  }

  function updateProjectsLoadIndicator() {
    const isLoadingProjects =
      state.paging.loading || state.paging.prefetchRunning || state.paging.inFlight.size > 0;

    if (isLoadingProjects) {
      const suffix = state.paging.loadedTo > 0 ? ` (${state.paging.loadedTo} loaded)` : '';
      setProjectsStatus(`Loading projects…${suffix}`, true);
    } else if (state.projects.length === 0) {
      setProjectsStatus('No projects on this page.', false);
    } else {
      const totalNote = formatProjectTotalsLabel();
      setProjectsStatus(`Select a project${totalNote}`, false);
    }

    updateSyncAllButton();
    updateRefreshButtons();
  }

  function setRefreshButtonLoading(buttonEl, iconEl, loading) {
    if (!buttonEl || !iconEl) return;
    buttonEl.disabled = loading;
    iconEl.classList.toggle('animate-spin', loading);
  }

  function updateRefreshButtons() {
    const projectsBusy =
      state.paging.loading || state.paging.prefetchRunning || state.syncingAll;
    els.refreshProjectsBtn.disabled = projectsBusy;

    const canRefreshFiles =
      Boolean(state.selected) &&
      !state.syncing &&
      !(state.selected && isProjectSyncingOnServer(state.selected.projectId));
    els.refreshFilesBtn.disabled = !canRefreshFiles;
  }

  async function refreshProjects() {
    if (state.paging.loading || state.syncingAll || state.paging.prefetchRunning) return;

    showError('');
    setRefreshButtonLoading(els.refreshProjectsBtn, els.refreshProjectsIcon, true);
    state.paging.pageCache.clear();
    state.paging.inFlight.clear();
    state.paging.loadedTo = 0;
    state.paging.allLoaded = false;
    state.paging.prefetchRunning = false;

    try {
      await loadProjects(state.paging.offset);
    } catch (error) {
      showError(error.message || 'Could not refresh projects.');
    } finally {
      setRefreshButtonLoading(els.refreshProjectsBtn, els.refreshProjectsIcon, false);
      updateRefreshButtons();
    }
  }

  async function refreshFiles() {
    if (!state.selected || state.syncing || isProjectSyncingOnServer(state.selected.projectId)) {
      return;
    }

    showError('');
    setRefreshButtonLoading(els.refreshFilesBtn, els.refreshFilesIcon, true);

    try {
      const project = state.selected;
      setFilesStatus(`Loading files for “${project.projectName}”…`, true);
      els.filesList.innerHTML = '';

      const response = await apiFetch(
        `/api/projects/${encodeURIComponent(project.projectId)}/documents`
      );
      const data = await response.json();
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to load files');
      }

      if (String(state.selected?.projectId) !== String(project.projectId)) return;

      state.documents = data.documents || [];
      renderFiles();
      updateSyncButton();
    } catch (error) {
      if (String(state.selected?.projectId) === String(project.projectId)) {
        setFilesStatus('Could not load files.', false);
        showError(error.message);
        updateSyncButton();
      }
    } finally {
      setRefreshButtonLoading(els.refreshFilesBtn, els.refreshFilesIcon, false);
      updateRefreshButtons();
    }
  }

  function formatProjectNumber(project) {
    if (project?.projectNumber != null && project.projectNumber !== '') {
      return String(project.projectNumber);
    }
    if (project?.projectId != null && project.projectId !== '') {
      return String(project.projectId);
    }
    return '—';
  }

  function renderProjects() {
    if (state.projects.length === 0) {
      els.projectsList.innerHTML = '';
      updateProjectsLoadIndicator();
      updateProjectsPager();
      updateRefreshButtons();
      return;
    }

    updateProjectsLoadIndicator();
    els.projectsList.innerHTML = state.projects
      .map((project, index) => {
        const selected =
          state.selected && String(state.selected.projectId) === String(project.projectId);
        const rowNumber = state.paging.offset + index + 1;
        const projectNumber = formatProjectNumber(project);
        const archived = isLikelyArchivedProject(project);
        return `
          <li>
            <button
              type="button"
              data-id="${escapeHtml(project.projectId)}"
              class="${archived ? 'relative overflow-hidden pr-10' : ''} flex w-full items-center gap-2 px-3 py-2.5 text-left text-sm transition ${
                selected
                  ? 'border-l-2 border-blue-600 bg-blue-50 font-medium text-blue-900'
                  : 'border-l-2 border-transparent text-gray-800 hover:bg-gray-50'
              } ${archived ? 'opacity-95' : ''}"
            >
              <span class="w-8 shrink-0 text-xs tabular-nums text-gray-400">${rowNumber}</span>
              <span class="shrink-0 rounded bg-gray-100 px-1.5 py-0.5 text-xs tabular-nums text-gray-600 ${
                selected ? 'bg-blue-100 text-blue-800' : ''
              }">${escapeHtml(projectNumber)}</span>
              <span class="min-w-0 flex-1 truncate">${escapeHtml(project.projectName)}</span>
              ${renderProjectNewBadge(project)}
              ${renderProjectStatusRibbon(project)}
            </button>
          </li>
        `;
      })
      .join('');

    els.projectsList.querySelectorAll('button[data-id]').forEach((button) => {
      button.addEventListener('click', () => {
        if (state.syncing || state.syncingAll) return;
        const project = state.projects.find(
          (item) => String(item.projectId) === String(button.dataset.id)
        );
        if (project) selectProject(project);
      });
    });
    updateProjectsPager();
    updateRefreshButtons();
  }

  function renderFiles() {
    if (!state.selected) {
      els.filesList.innerHTML = '';
      setFilesStatus('Select a project to view files', false);
      return;
    }

    if (state.documents.length === 0) {
      els.filesList.innerHTML = '';
      setFilesStatus(`No files in “${state.selected.projectName}”`, false);
      return;
    }

    const projectNumber = formatProjectNumber(state.selected);
    setFilesStatus(
      `${state.documents.length} file(s) in “${state.selected.projectName}” (#${projectNumber})`,
      false
    );
    els.filesList.innerHTML = state.documents
      .map(
        (doc, index) => `
          <li class="flex items-center gap-2 px-3 py-2 text-sm text-gray-800">
            <span class="w-8 shrink-0 text-xs tabular-nums text-gray-400">${index + 1}</span>
            <span class="min-w-0 flex-1 truncate" title="${escapeHtml(doc.filename)}">${escapeHtml(doc.filename)}</span>
          </li>
        `
      )
      .join('');
    updateRefreshButtons();
  }

  function getChunkOffset(offset) {
    return Math.floor(offset / state.paging.fetchSize) * state.paging.fetchSize;
  }

  function getPageFromChunk(offset) {
    const chunkOffset = getChunkOffset(offset);
    const chunk = state.paging.pageCache.get(chunkOffset);
    if (!chunk) return null;

    const indexInChunk = offset - chunkOffset;
    const projects = chunk.projects.slice(indexInChunk, indexInChunk + state.paging.pageSize);
    return {
      projects,
      hasMore: chunk.hasMore || indexInChunk + projects.length < chunk.projects.length,
      chunkOffset,
    };
  }

  function canGoNextPage() {
    const nextOffset = state.paging.offset + state.paging.pageSize;
    const currentChunk = state.paging.pageCache.get(getChunkOffset(state.paging.offset));
    const nextPage = getPageFromChunk(nextOffset);
    if (nextPage && nextPage.projects.length > 0) return true;
    if (state.paging.inFlight.has(getChunkOffset(nextOffset))) return true;
    return Boolean(currentChunk?.hasMore);
  }

  async function fetchProjectsPage(offset) {
    const normalizedOffset = Math.max(0, offset);
    if (state.paging.pageCache.has(normalizedOffset)) {
      return state.paging.pageCache.get(normalizedOffset);
    }

    if (state.paging.inFlight.has(normalizedOffset)) {
      return state.paging.inFlight.get(normalizedOffset);
    }

    const fetchPromise = (async () => {
      const params = new URLSearchParams({
        offset: String(normalizedOffset),
        limit: String(state.paging.fetchSize),
      });
      const response = await apiFetch(`/api/projects?${params.toString()}`);
      const data = await response.json();
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to load projects');
      }

      const page = {
        offset: normalizedOffset,
        projects: (data.projects || []).map((project) => {
          const alreadyUploaded =
            Boolean(project.alreadyUploaded) ||
            state.uploadedProjectIds.has(String(project.projectId));
          return {
            ...project,
            alreadyUploaded,
            isNew: !alreadyUploaded,
          };
        }),
        loadedTo: parseNonNegativeInt(
          data.loadedTo,
          normalizedOffset + (data.projects || []).length
        ),
        hasMore: Boolean(data.hasMore),
      };
      if (typeof data.uploadedProjectCount === 'number') {
        // Keep local set in sync when API returns flags.
        for (const project of page.projects) {
          if (project.alreadyUploaded) {
            state.uploadedProjectIds.add(String(project.projectId));
          }
        }
      }
      state.paging.pageCache.set(normalizedOffset, page);
      state.paging.loadedTo = Math.max(state.paging.loadedTo, page.loadedTo);
      updateProjectsLoadIndicator();
      return page;
    })();

    state.paging.inFlight.set(normalizedOffset, fetchPromise);
    try {
      return await fetchPromise;
    } finally {
      state.paging.inFlight.delete(normalizedOffset);
    }
  }

  async function prefetchRemainingPages() {
    if (state.paging.allLoaded) return;
    if (state.paging.prefetchRunning) return;
    state.paging.prefetchRunning = true;
    state.paging.allLoaded = false;
    updateProjectsLoadIndicator();

    try {
      let offset = 0;
      while (true) {
        const page = state.paging.pageCache.get(offset) || (await fetchProjectsPage(offset));
        if (!page.hasMore) break;
        offset += state.paging.fetchSize;

        // Yield so UI stays responsive while prefetching large project lists.
        await new Promise((resolve) => setTimeout(resolve, 0));
      }

      if (state.projects.length > 0) {
        state.paging.allLoaded = true;
        renderProjects();
      } else {
        state.paging.allLoaded = true;
      }
    } catch {
      // Background prefetch is best-effort; no UI interruption needed.
      state.paging.allLoaded = false;
    } finally {
      state.paging.prefetchRunning = false;
      updateProjectsLoadIndicator();
    }
  }

  async function loadProjects(offset = 0) {
    showError('');
    state.paging.loading = true;
    state.paging.offset = Math.max(0, offset);
    updateProjectsLoadIndicator();
    updateProjectsPager();

    const viewOffset = state.paging.offset;
    const chunkOffset = getChunkOffset(viewOffset);
    const cachedView = getPageFromChunk(viewOffset);
    if (!cachedView) {
      setProjectsStatus('Loading projects…', true);
      els.projectsList.innerHTML = '';
    }

    try {
      const chunk = state.paging.pageCache.get(chunkOffset) || (await fetchProjectsPage(chunkOffset));
      const view = cachedView || getPageFromChunk(viewOffset);

      state.projects = view?.projects || [];
      state.paging.loadedTo = Math.max(state.paging.loadedTo, chunk.loadedTo);
      state.paging.hasMore = Boolean(chunk.hasMore);

      renderProjects();

      // Start background prefetch once first page is displayed.
      prefetchRemainingPages();
    } catch (error) {
      setProjectsStatus('Could not load projects.', false);
      showError(error.message);
    } finally {
      state.paging.loading = false;
      updateProjectsLoadIndicator();
      updateProjectsPager();
    }
  }

  async function selectProject(project) {
    state.selected = project;
    state.documents = [];
    renderProjects();
    renderFiles();
    updateSyncButton();
    updateRefreshButtons();

    setFilesStatus(`Loading files for “${project.projectName}”…`, true);
    els.filesList.innerHTML = '';
    showError('');

    try {
      const response = await apiFetch(
        `/api/projects/${encodeURIComponent(project.projectId)}/documents`
      );
      const data = await response.json();
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to load files');
      }

      // Ignore stale responses if user clicked another project
      if (String(state.selected?.projectId) !== String(project.projectId)) return;

      state.documents = data.documents || [];
      renderFiles();
      updateSyncButton();
    } catch (error) {
      if (String(state.selected?.projectId) !== String(project.projectId)) return;
      setFilesStatus('Could not load files.', false);
      showError(error.message);
      updateSyncButton();
    }
  }

  function formatUploadBytes(bytes) {
    const value = Math.max(0, Number(bytes) || 0);
    if (value >= 1024 * 1024 * 1024) {
      return `${(value / (1024 * 1024 * 1024)).toFixed(2)} GB`;
    }
    if (value >= 1024 * 1024) {
      return `${(value / (1024 * 1024)).toFixed(1)} MB`;
    }
    if (value >= 1024) {
      return `${(value / 1024).toFixed(1)} KB`;
    }
    return `${Math.round(value)} B`;
  }

  function formatUploadSizeLabel(bytesDone, bytesTotal) {
    const done = Math.max(0, Number(bytesDone) || 0);
    const total = Math.max(0, Number(bytesTotal) || 0);
    if (total > 0) {
      return `${formatUploadBytes(done)} / ${formatUploadBytes(total)}`;
    }
    return formatUploadBytes(done);
  }

  const activeTransferToasts = new Map();
  let lastSyncAllUiUpdateAt = 0;
  let syncAllUiUpdateTimer = null;
  let syncAllUiForcePending = false;

  const TRANSFER_TOAST_ICONS = {
    downloading: `
      <svg class="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
        <path stroke-linecap="round" stroke-linejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
      </svg>
    `,
    uploading: `
      <svg class="h-4 w-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
        <path stroke-linecap="round" stroke-linejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" />
      </svg>
    `,
  };

  function getTransferToastKey(data = {}) {
    const projectId = data.projectId != null ? String(data.projectId) : 'project';
    const documentId = data.documentId != null ? String(data.documentId) : data.filename || 'file';
    return `${projectId}:${documentId}`;
  }

  function updateTransferToastHeader() {
    if (!els.transferToastHeader) return;
    let downloading = 0;
    let uploading = 0;
    for (const toast of activeTransferToasts.values()) {
      if (toast.stage === 'downloading') downloading += 1;
      else uploading += 1;
    }
    const total = downloading + uploading;
    els.transferToastHeader.classList.toggle('is-visible', total > 0);
    if (els.transferToastDownloadCount) {
      els.transferToastDownloadCount.textContent = String(downloading);
    }
    if (els.transferToastUploadCount) {
      els.transferToastUploadCount.textContent = String(uploading);
    }
    if (els.transferToastHeader.querySelector) {
      const title = $('transfer-toast-header-title');
      if (title) {
        title.textContent =
          total === 0
            ? 'Active transfers'
            : `${total} transfer${total === 1 ? '' : 's'} in progress`;
      }
    }
  }

  function ensureTransferToast(key) {
    let toast = activeTransferToasts.get(key);
    if (toast) return toast;

    const el = document.createElement('div');
    el.className = 'transfer-toast transfer-toast-download';
    el.innerHTML = `
      <div class="flex items-start gap-3">
        <div data-role="icon" class="transfer-toast-icon">${TRANSFER_TOAST_ICONS.downloading}</div>
        <div class="min-w-0 flex-1">
          <div class="flex items-start justify-between gap-2">
            <div class="min-w-0">
              <div data-role="stage" class="transfer-toast-stage text-[11px] font-semibold uppercase tracking-[0.08em]">Downloading</div>
              <div data-role="filename" class="mt-0.5 truncate text-sm font-semibold text-slate-900"></div>
              <div data-role="project" class="mt-0.5 truncate text-[11px] text-slate-500"></div>
            </div>
            <div data-role="percent" class="transfer-toast-percent shrink-0 text-sm font-bold tabular-nums">0%</div>
          </div>
          <div data-role="track" class="transfer-toast-track" aria-hidden="true">
            <div data-role="bar" class="transfer-toast-bar"></div>
          </div>
          <div class="mt-1.5 flex items-center justify-between gap-2">
            <div data-role="size" class="transfer-toast-size text-[11px] font-medium tabular-nums">0 B</div>
            <div data-role="hint" class="text-[10px] font-medium uppercase tracking-wide text-slate-400">In progress</div>
          </div>
        </div>
      </div>
    `;
    els.transferToastStack.appendChild(el);
    toast = {
      el,
      iconEl: el.querySelector('[data-role="icon"]'),
      stageEl: el.querySelector('[data-role="stage"]'),
      filenameEl: el.querySelector('[data-role="filename"]'),
      projectEl: el.querySelector('[data-role="project"]'),
      percentEl: el.querySelector('[data-role="percent"]'),
      barEl: el.querySelector('[data-role="bar"]'),
      sizeEl: el.querySelector('[data-role="size"]'),
      stage: null,
      lastUiAt: 0,
      pending: null,
      flushTimer: null,
    };
    activeTransferToasts.set(key, toast);
    updateTransferToastHeader();
    return toast;
  }

  function applyTransferToastStage(toast, stage) {
    if (toast.stage === stage) return;
    toast.stage = stage;
    const isDownload = stage === 'downloading';
    toast.el.classList.toggle('transfer-toast-download', isDownload);
    toast.el.classList.toggle('transfer-toast-upload', !isDownload);
    toast.iconEl.innerHTML = TRANSFER_TOAST_ICONS[stage] || TRANSFER_TOAST_ICONS.uploading;
    toast.stageEl.textContent = isDownload ? 'Downloading' : 'Uploading';
    updateTransferToastHeader();
  }

  function upsertTransferToast(data = {}) {
    if (!els.transferToastStack) return;
    const key = getTransferToastKey(data);
    const toast = ensureTransferToast(key);
    const stage = data.stage === 'downloading' ? 'downloading' : 'uploading';
    const percent = Math.max(0, Math.min(100, Number(data.percent) || 0));
    const bytesDone =
      Number(data.bytesDone) ||
      (stage === 'downloading' ? Number(data.bytesDownloaded) : Number(data.bytesUploaded)) ||
      0;
    const bytesTotal = Number(data.bytesTotal) || 0;
    const force = percent >= 100 || percent === 0 || toast.stage !== stage;
    const now = Date.now();

    if (!force && toast.lastUiAt && now - toast.lastUiAt < TRANSFER_TOAST_MIN_INTERVAL_MS) {
      toast.pending = { stage, percent, bytesDone, bytesTotal, data };
      if (!toast.flushTimer) {
        toast.flushTimer = window.setTimeout(() => {
          toast.flushTimer = null;
          const pending = toast.pending;
          toast.pending = null;
          if (!pending) return;
          upsertTransferToast({
            ...pending.data,
            stage: pending.stage,
            percent: pending.percent,
            bytesDone: pending.bytesDone,
            bytesTotal: pending.bytesTotal,
            bytesDownloaded: pending.stage === 'downloading' ? pending.bytesDone : undefined,
            bytesUploaded: pending.stage === 'uploading' ? pending.bytesDone : undefined,
          });
        }, TRANSFER_TOAST_MIN_INTERVAL_MS);
      }
      return;
    }

    toast.lastUiAt = now;
    toast.pending = null;
    applyTransferToastStage(toast, stage);
    toast.filenameEl.textContent = data.filename || 'file';
    toast.filenameEl.title = data.filename || 'file';
    toast.projectEl.textContent = data.projectName || '';
    toast.projectEl.classList.toggle('hidden', !data.projectName);
    toast.percentEl.textContent = `${percent}%`;
    toast.barEl.style.width = `${percent}%`;
    toast.sizeEl.textContent = formatUploadSizeLabel(bytesDone, bytesTotal);
  }

  function removeTransferToast(data = {}) {
    const key = getTransferToastKey(data);
    const toast = activeTransferToasts.get(key);
    if (!toast) return;
    if (toast.flushTimer) {
      window.clearTimeout(toast.flushTimer);
      toast.flushTimer = null;
    }
    toast.el.classList.add('is-leaving');
    window.setTimeout(() => {
      toast.el.remove();
      activeTransferToasts.delete(key);
      updateTransferToastHeader();
    }, 160);
  }

  function clearAllTransferToasts() {
    for (const toast of activeTransferToasts.values()) {
      if (toast.flushTimer) {
        window.clearTimeout(toast.flushTimer);
      }
      toast.el.remove();
    }
    activeTransferToasts.clear();
    updateTransferToastHeader();
  }

  function setProgress(percent, text) {
    const clamped = Math.max(0, Math.min(100, Number(percent) || 0));
    els.progressBar.style.width = `${clamped}%`;
    els.progressPercent.textContent = `${clamped}%`;
    if (text) els.progressText.textContent = text;
  }

  function setStats(total, ok, fail) {
    els.statTotal.textContent = String(total);
    els.statOk.textContent = String(ok);
    els.statFail.textContent = String(fail);
  }

  function updateUploadProgress(total, ok, fail) {
    const safeTotal = Number(total) || 0;
    const uploaded = Number(ok) || 0;
    const failed = Number(fail) || 0;
    const percent =
      safeTotal === 0 ? 100 : Math.round((uploaded / safeTotal) * 100);

    setStats(safeTotal, uploaded, failed);
    setProgress(percent, `${uploaded} / ${safeTotal} files`);
  }

  function addArchivedSkipResult(data) {
    const li = document.createElement('li');
    li.className = 'rounded bg-amber-50 px-2 py-1 text-amber-900';
    li.textContent = formatArchivedProjectSkipLine(null, data);
    els.resultList.prepend(li);
    while (els.resultList.children.length > MAX_SINGLE_RESULT_ROWS) {
      els.resultList.removeChild(els.resultList.lastElementChild);
    }
  }

  function addResult(filename, ok, detail) {
    const li = document.createElement('li');
    li.className = ok
      ? 'rounded bg-green-50 px-2 py-1 text-green-800'
      : 'rounded bg-red-50 px-2 py-1 text-red-800';
    li.innerHTML = ok
      ? `✓ ${escapeHtml(filename)}`
      : `✗ ${escapeHtml(filename)} — ${escapeHtml(detail || 'failed')}`;
    els.resultList.prepend(li);
    while (els.resultList.children.length > MAX_SINGLE_RESULT_ROWS) {
      els.resultList.removeChild(els.resultList.lastElementChild);
    }
  }

  function updateSyncAllProgressUi(options = {}) {
    const force = Boolean(options.force);
    const now = Date.now();
    if (!force && now - lastSyncAllUiUpdateAt < SYNC_ALL_UI_MIN_INTERVAL_MS) {
      syncAllUiForcePending = syncAllUiForcePending || Boolean(options.forceNext);
      if (!syncAllUiUpdateTimer) {
        syncAllUiUpdateTimer = window.setTimeout(() => {
          syncAllUiUpdateTimer = null;
          updateSyncAllProgressUi({ force: true });
        }, SYNC_ALL_UI_MIN_INTERVAL_MS);
      }
      return;
    }
    lastSyncAllUiUpdateAt = now;
    if (syncAllUiUpdateTimer) {
      window.clearTimeout(syncAllUiUpdateTimer);
      syncAllUiUpdateTimer = null;
    }
    syncAllUiForcePending = false;

    const {
      projectsTotal,
      projectsDone,
      filesOk,
      filesSkipped,
      filesFail,
      projectsSkippedArchived,
      currentProjectName,
      paused,
    } = state.syncAll;
    const percent =
      projectsTotal === 0 ? 100 : Math.round((projectsDone / projectsTotal) * 100);

    els.syncAllProjectsDone.textContent = String(projectsDone);
    els.syncAllProjectsTotal.textContent = String(projectsTotal);
    els.syncAllFilesOk.textContent = String(filesOk);
    els.syncAllFilesSkipped.textContent = String(filesSkipped || 0);
    els.syncAllProjectsSkippedArchived.textContent = String(projectsSkippedArchived || 0);
    els.syncAllFilesFail.textContent = String(filesFail);
    els.syncAllProgressBar.style.width = `${percent}%`;
    els.syncAllProgressPercent.textContent = `${percent}%`;
    updateSyncAllElapsedLabel();

    const archivedNote =
      projectsSkippedArchived > 0 ? ` · ${projectsSkippedArchived} archived skipped` : '';

    if (state.syncingAll && paused) {
      els.syncAllProgressText.textContent = state.syncAll.pausedDueToCrash
        ? `Paused (backend issue) · ${projectsDone} / ${projectsTotal} projects${archivedNote}`
        : `Paused · ${projectsDone} / ${projectsTotal} projects${archivedNote}`;
    } else if (state.syncingAll && currentProjectName) {
      els.syncAllProgressText.textContent = `${projectsDone} / ${projectsTotal} projects${archivedNote} · ${currentProjectName}`;
    } else if (state.syncingAll) {
      els.syncAllProgressText.textContent = `${projectsDone} / ${projectsTotal} projects${archivedNote}`;
    } else {
      els.syncAllProgressText.textContent = `${projectsDone} / ${projectsTotal} projects done${archivedNote}`;
    }
  }

  function addSyncAllHistory(message, type = 'info') {
    const li = document.createElement('li');
    const styles = {
      success: 'rounded bg-green-50 px-2 py-1 text-green-800',
      error: 'rounded bg-red-50 px-2 py-1 text-red-800',
      info: 'rounded bg-gray-50 px-2 py-1 text-gray-700',
      archived: 'rounded bg-amber-50 px-2 py-1 text-amber-900',
    };
    li.className = styles[type] || styles.info;
    li.textContent = message;
    els.syncAllResultList.appendChild(li);
    while (els.syncAllResultList.children.length > MAX_SYNC_ALL_HISTORY_ROWS) {
      els.syncAllResultList.removeChild(els.syncAllResultList.firstElementChild);
    }
    // Avoid forcing scroll on every row — only when near bottom.
    const list = els.syncAllResultList;
    const nearBottom = list.scrollHeight - list.scrollTop - list.clientHeight < 80;
    if (nearBottom) {
      list.scrollTop = list.scrollHeight;
    }
  }

  function readSseStream(response, onEvent) {
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let streamCompleted = false;

    return (async function pump() {
      try {
        while (true) {
          const { done, value } = await reader.read();
          if (done) break;

          buffer += decoder.decode(value, { stream: true });
          const chunks = buffer.split('\n\n');
          buffer = chunks.pop() || '';

          for (const chunk of chunks) {
            if (!chunk.trim()) continue;

            let event = 'message';
            const dataLines = [];

            for (const line of chunk.split('\n')) {
              if (line.startsWith('event:')) event = line.slice(6).trim();
              else if (line.startsWith('data:')) dataLines.push(line.slice(5).trim());
            }

            if (!dataLines.length) continue;
            try {
              if (event === 'complete') {
                streamCompleted = true;
              }
              onEvent(event, JSON.parse(dataLines.join('\n')));
            } catch {
              // ignore bad chunks
            }
          }
        }
      } catch (error) {
        if (!streamCompleted) {
          throw new Error(`Backend connection lost while receiving progress: ${error.message}`);
        }
        throw error;
      }

      if (!streamCompleted) {
        throw new Error('Backend connection lost before sync completed.');
      }
    })();
  }

  async function syncProjectToSharePoint(project, handlers = {}) {
    const {
      onStarted,
      onFileSuccess,
      onFileError,
      onFileTransferProgress,
      onComplete,
      onError,
      onStatus,
    } = handlers;

    let response;
    try {
      response = await apiFetch(
        `/api/projects/${encodeURIComponent(project.projectId)}/sync`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ projectName: project.projectName }),
        }
      );
    } catch (error) {
      throw new Error(`Backend connection lost before sync started: ${error.message}`);
    }

    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      throw new Error(body.error || `Sync failed (${response.status})`);
    }

    let summary = {
      succeeded: 0,
      failed: 0,
      total: 0,
      success: false,
      fatalError: null,
    };

    await readSseStream(response, (event, data) => {
      if (event === 'status' && onStatus) onStatus(data);
      if (event === 'started') {
        summary.total = data.total || 0;
        if (onStarted) onStarted(data);
      }
      if (event === 'file-transfer-progress' && onFileTransferProgress) onFileTransferProgress(data);
      if (event === 'file-upload-progress' && onFileTransferProgress) onFileTransferProgress(data);
      if (event === 'file-success' && onFileSuccess) onFileSuccess(data);
      if (event === 'file-error' && onFileError) onFileError(data);
      if (event === 'complete') {
        summary = {
          succeeded: data.succeeded || 0,
          failed: data.failed || 0,
          total: data.total || 0,
          success: Boolean(data.success),
          skippedArchivedProject: Boolean(data.skippedArchivedProject),
          projectId: data.projectId || null,
          projectName: data.projectName || null,
          projectNumber: data.projectNumber || null,
          phaseName: data.phaseName || null,
          message: data.message || null,
          counts: data.counts || null,
          fatalError: null,
        };
        if (onComplete) onComplete(data);
      }
      if (event === 'error') {
        summary.fatalError = data.error || 'Sync failed';
        if (onError) onError(data);
      }
    });

    if (summary.fatalError) {
      throw new Error(summary.fatalError);
    }

    return summary;
  }

  async function startSync() {
    if (!state.selected || state.syncing || (state.syncingAll && !state.syncAll.paused)) {
      return;
    }
    if (isUploadBlockedBySchedule()) {
      showError('Uploads are disabled while the scheduled sync is running (about 2–3 hours).');
      return;
    }
    const isResume = state.singleSync.pausedDueToCrash;
    if (!isResume && state.documents.length === 0) {
      return;
    }

    state.syncing = true;
    state.singleSync.pausedDueToCrash = false;
    state.singleSync.pauseReason = '';
    hidePauseNotice();
    updateSyncButton();
    updateSyncAllButton();
    showError('');

    els.progressBox.classList.remove('hidden');
    if (!isResume) {
      els.resultList.innerHTML = '';
      updateUploadProgress(state.documents.length, 0, 0);
      clearAllTransferToasts();
    }

    try {
      await syncProjectToSharePoint(state.selected, {
        onStarted: (data) => updateUploadProgress(data.total || 0, 0, 0),
        onStatus: (data) => {
          if (data?.stage === 'skipped-archived' || data?.stage === 'checking-project') {
            setProgress(0, data.message || 'Checking project…');
          }
        },
        onFileTransferProgress: (data) => {
          upsertTransferToast({
            ...data,
            projectId: data.projectId || state.selected?.projectId,
            projectName: data.projectName || state.selected?.projectName,
          });
        },
        onFileSuccess: (data) => {
          removeTransferToast({
            ...data,
            projectId: data.projectId || state.selected?.projectId,
          });
          if (
            data?.skippedAlreadyUploaded ||
            data?.skippedNoExtension ||
            data?.skippedNameConflict ||
            data?.skippedDuplicateFilename
          ) {
            addResult(data.filename, true, data.message || 'Skipped');
          } else {
            addResult(data.filename, true);
          }
          updateUploadProgress(data.total, data.succeeded, data.failed);
        },
        onFileError: (data) => {
          removeTransferToast({
            ...data,
            projectId: data.projectId || state.selected?.projectId,
          });
          addResult(data.filename, false, data.error);
          updateUploadProgress(data.total, data.succeeded, data.failed);
        },
        onComplete: (data) => {
          clearAllTransferToasts();
          updateUploadProgress(data.total || 0, data.succeeded || 0, data.failed || 0);
          if (data.skippedArchivedProject) {
            setProgress(100, data.message || 'Skipped archived project');
            addArchivedSkipResult(data);
          } else if (!data.skippedArchivedProject) {
            markProjectUploadedLocally(data.projectId || state.selected?.projectId);
            renderProjects();
            updateSyncAllControls();
          }
        },
        onError: (data) => {
          clearAllTransferToasts();
          setProgress(0, data.error || 'Sync failed');
          showError(data.error || 'Sync failed');
        },
      });
    } catch (error) {
      if (isBackendConnectionError(error)) {
        const message = getBackendPauseMessage(error);
        state.singleSync.pausedDueToCrash = true;
        state.singleSync.pauseReason = message;
        showPauseNotice(message);
        setProgress(0, 'Paused — backend connection lost');
        showError('');
      } else {
        setProgress(0, 'Error');
        showError(error.message);
      }
    } finally {
      clearAllTransferToasts();
      state.syncing = false;
      updateSyncButton();
      updateSyncAllButton();
      renderProjects();
    }
  }

  async function continueSyncAll(projects) {
    if (!projects.length) {
      finishSyncAll();
      return;
    }
    if (state.syncAll.processing) return;

    state.syncAll.processing = true;
    state.syncAll.remainingProjects = projects;
    saveSyncAllSnapshot();
    const activeProjects = new Set();
    let nextIndex = 0;

    function updateActiveProjectSummary() {
      state.syncAll.activeProjects = [...activeProjects];
      if (activeProjects.size === 0) {
        state.syncAll.currentProjectName = '';
      } else if (activeProjects.size === 1) {
        state.syncAll.currentProjectName = state.syncAll.activeProjects[0];
      } else {
        state.syncAll.currentProjectName = `${activeProjects.size} projects syncing in parallel`;
      }
      updateSyncAllControls();
      updateSyncAllProgressUi();
    }

    function takeNextProject() {
      if (nextIndex >= projects.length) {
        state.syncAll.remainingProjects = [];
        return null;
      }
      const project = projects[nextIndex];
      nextIndex += 1;
      state.syncAll.remainingProjects = projects.slice(nextIndex);
      return project;
    }

    try {
      const workerCount = Math.min(SYNC_ALL_PROJECT_CONCURRENCY, projects.length);
      const workers = Array.from({ length: workerCount }, () =>
        (async () => {
          while (true) {
            await waitIfPaused();
            if (state.syncAll.pausedDueToCrash) return;

            const project = takeNextProject();
            if (!project) return;
            saveSyncAllSnapshot();

            const projectLabel = project.projectName || `Project ${project.projectId}`;
            activeProjects.add(projectLabel);
            updateActiveProjectSummary();

            try {
              // Known archived from project list — skip all Filevine/SharePoint calls.
              if (isLikelyArchivedProject(project)) {
                state.syncAll.projectsDone += 1;
                state.syncAll.projectsSkippedArchived += 1;
                updateSyncAllProgressUi({ force: true });
                saveSyncAllSnapshot();
                addSyncAllHistory(formatArchivedProjectSkipLine(project), 'archived');
                continue;
              }

              const summary = await syncProjectToSharePoint(project, {
                onFileTransferProgress: (data) => {
                  upsertTransferToast({
                    ...data,
                    projectId: data.projectId || project.projectId,
                    projectName: data.projectName || project.projectName,
                  });
                },
                onFileSuccess: (data) => {
                  if (
                    data?.skippedAlreadyUploaded ||
                    data?.skippedNoExtension ||
                    data?.skippedNameConflict ||
                    data?.skippedDuplicateFilename
                  ) {
                    state.syncAll.filesSkipped += 1;
                    removeTransferToast({
                      ...data,
                      projectId: data.projectId || project.projectId,
                    });
                    updateSyncAllProgressUi();
                    return;
                  }
                  state.syncAll.filesOk += 1;
                  removeTransferToast({
                    ...data,
                    projectId: data.projectId || project.projectId,
                  });
                  updateSyncAllProgressUi();
                },
                onFileError: (data) => {
                  state.syncAll.filesFail += 1;
                  removeTransferToast({
                    ...data,
                    projectId: data.projectId || project.projectId,
                  });
                  updateSyncAllProgressUi({ force: true });
                  addSyncAllHistory(
                    `✗ ${data.filename} — ${data.error || 'failed'} (${project.projectName})`,
                    'error'
                  );
                },
              });

              state.syncAll.projectsDone += 1;
              updateSyncAllProgressUi({ force: true });
              saveSyncAllSnapshot({ force: true });
              if (summary.skippedArchivedProject) {
                state.syncAll.projectsSkippedArchived += 1;
                updateSyncAllProgressUi({ force: true });
                addSyncAllHistory(formatArchivedProjectSkipLine(project, summary), 'archived');
              } else {
                markProjectUploadedLocally(project.projectId);
                const uploaded = Number(summary.counts?.newlyUploaded);
                const skipped = Number(summary.counts?.skippedAlreadyUploaded);
                const failed = Number(summary.counts?.failed ?? summary.failed) || 0;
                const uploadedCount = Number.isFinite(uploaded) ? uploaded : Number(summary.succeeded) || 0;
                const skippedCount = Number.isFinite(skipped) ? skipped : 0;
                addSyncAllHistory(
                  `Done: ${project.projectName} · ${uploadedCount} uploaded, ${failed} failed, ${skippedCount} skipped`,
                  failed > 0 ? 'error' : 'success'
                );
              }
            } catch (error) {
              if (isBackendConnectionError(error) || isSharePointConfigError(error)) {
                if (!state.syncAll.pausedDueToCrash) {
                  const remaining = [project, ...projects.slice(nextIndex)];
                  pauseSyncAllDueToCrash(error, remaining);
                }
                return;
              }

              state.syncAll.projectsDone += 1;
              updateSyncAllProgressUi();
              saveSyncAllSnapshot();
              addSyncAllHistory(
                `Failed project: ${project.projectName} — ${error.message}`,
                'error'
              );
            } finally {
              activeProjects.delete(projectLabel);
              updateActiveProjectSummary();
            }
          }
        })()
      );

      await Promise.all(workers);
      if (state.syncAll.pausedDueToCrash) return;

      finishSyncAll();
    } finally {
      state.syncAll.processing = false;
    }
  }

  function finishSyncAll() {
    stopSyncAllTimer();
    state.syncAll.currentProjectName = '';
    state.syncAll.activeProjects = [];
    state.syncAll.paused = false;
    state.syncAll.pausedDueToCrash = false;
    state.syncAll.pauseReason = '';
    state.syncAll.remainingProjects = [];
    state.syncAll.resumeSnapshot = null;
    state.syncAll.elapsedMs = 0;
    if (state.syncAll.pauseResolver) {
      state.syncAll.pauseResolver();
      state.syncAll.pauseResolver = null;
    }
    state.syncingAll = false;
    hidePauseNotice();
    clearAllTransferToasts();
    clearSyncAllSnapshot();
    updateSyncButton();
    updateSyncAllControls();
    updateProjectsPager();
    updateSyncAllProgressUi();
    renderProjects();
    addSyncAllHistory(
      `Finished all projects · ${state.syncAll.filesOk} uploaded, ${state.syncAll.filesSkipped} files skipped, ${state.syncAll.projectsSkippedArchived} archived projects skipped, ${state.syncAll.filesFail} failed`,
      state.syncAll.filesFail > 0
        ? 'error'
        : (state.syncAll.filesSkipped > 0 || state.syncAll.projectsSkippedArchived > 0 ? 'info' : 'success')
    );
  }

  async function startSyncAll(options = {}) {
    if (state.syncing || (state.syncingAll && !state.syncAll.paused)) return;
    if (isUploadBlockedBySchedule()) {
      showError('Uploads are disabled while the scheduled sync is running (about 2–3 hours).');
      return;
    }
    if (!state.paging.allLoaded) return;

    const onlyNew = Boolean(options.onlyNew);
    const allProjects = getAllCachedProjects();
    const sourceProjects = onlyNew ? getNewProjects() : allProjects;
    if (sourceProjects.length === 0) {
      showError(onlyNew ? 'No new projects available to sync.' : 'No projects available to sync.');
      return;
    }

    const isResume = state.syncingAll && state.syncAll.pausedDueToCrash;
    if (isResume) {
      resumeSyncAll();
      return;
    }

    const projectsToSync = sourceProjects;

    state.syncingAll = true;
    state.syncAll = {
      projectsTotal: sourceProjects.length,
      projectsDone: 0,
      filesOk: 0,
      filesSkipped: 0,
      filesFail: 0,
      projectsSkippedArchived: 0,
      elapsedMs: 0,
      timerStartedAt: null,
      timerIntervalId: null,
      currentProjectName: '',
      activeProjects: [],
      paused: false,
      pausedDueToCrash: false,
      pauseReason: '',
      pauseResolver: null,
      remainingProjects: projectsToSync,
      processing: false,
      resumeSnapshot: null,
    };
    clearSyncAllSnapshot();

    updateSyncButton();
    updateSyncAllControls();
    updateProjectsPager();
    showError('');
    hidePauseNotice();

    els.syncAllProgressBox.classList.remove('hidden');
    els.syncAllResultList.innerHTML = '';
    updateSyncAllProgressUi();
    startSyncAllTimer();
    if (onlyNew) {
      addSyncAllHistory(
        `Starting sync for ${projectsToSync.length} new project(s)…`,
        'info'
      );
    } else {
      addSyncAllHistory(`Starting sync for ${allProjects.length} project(s)…`, 'info');
    }

    await continueSyncAll(projectsToSync);
  }

  async function startSyncNewProjects() {
    await startSyncAll({ onlyNew: true });
  }

  async function resumeLastSyncAllRun() {
    if (state.syncing || state.syncingAll || isUploadBlockedBySchedule()) return;
    const snapshot = getResumeSnapshot();
    if (!snapshot) {
      showError('No interrupted sync run found to resume.');
      updateSyncAllControls();
      return;
    }

    const remainingProjects = resolveSnapshotRemainingProjects(snapshot);
    if (!remainingProjects.length) {
      showError('Could not restore remaining projects from the saved run.');
      updateSyncAllControls();
      return;
    }

    state.syncingAll = true;
    state.syncAll = {
      projectsTotal: Number(snapshot.projectsTotal) || remainingProjects.length + (Number(snapshot.projectsDone) || 0),
      projectsDone: Number(snapshot.projectsDone) || 0,
      filesOk: Number(snapshot.filesOk) || 0,
      filesSkipped: Number(snapshot.filesSkipped) || 0,
      filesFail: Number(snapshot.filesFail) || 0,
      projectsSkippedArchived: Number(snapshot.projectsSkippedArchived) || 0,
      elapsedMs: Number(snapshot.elapsedMs) || 0,
      timerStartedAt: null,
      timerIntervalId: null,
      currentProjectName: '',
      activeProjects: [],
      paused: false,
      pausedDueToCrash: false,
      pauseReason: '',
      pauseResolver: null,
      remainingProjects,
      processing: false,
      resumeSnapshot: snapshot,
    };

    updateSyncButton();
    updateSyncAllControls();
    updateProjectsPager();
    showError('');
    hidePauseNotice();
    els.syncAllProgressBox.classList.remove('hidden');
    updateSyncAllProgressUi();
    startSyncAllTimer();
    addSyncAllHistory(
      `Resuming previous run from ${remainingProjects.length} remaining project(s) at ${snapshot.progressPercent ?? 0}%…`,
      'info'
    );

    await continueSyncAll(remainingProjects);
  }

  els.syncBtn.addEventListener('click', startSync);
  els.syncAllBtn.addEventListener('click', () => startSyncAll());
  if (els.syncNewBtn) {
    els.syncNewBtn.addEventListener('click', startSyncNewProjects);
  }
  els.pauseSyncAllBtn.addEventListener('click', pauseSyncAll);
  els.resumeSyncAllBtn.addEventListener('click', resumeSyncAll);
  els.resumeLastSyncAllBtn.addEventListener('click', resumeLastSyncAllRun);
  els.loginForm.addEventListener('submit', (event) => {
    event.preventDefault();
    handleLogin(false);
  });
  els.sessionConflictCancel.addEventListener('click', hideSessionConflictModal);
  els.sessionConflictConfirm.addEventListener('click', () => {
    hideSessionConflictModal();
    handleLogin(true);
  });
  els.settingsBtn.addEventListener('click', openSettingsModal);
  els.scheduleBtn.addEventListener('click', openScheduleModal);
  els.scheduleStatusBar.addEventListener('click', openScheduleModal);
  els.scheduleCancelBtn.addEventListener('click', closeScheduleModal);
  els.scheduleForm.addEventListener('submit', saveSchedule);
  els.scheduleFrequencyDaily.addEventListener('change', updateScheduleFrequencyUi);
  els.scheduleFrequencyWeekly.addEventListener('change', updateScheduleFrequencyUi);
  els.scheduleModal.addEventListener('click', (event) => {
    if (event.target === els.scheduleModal) {
      closeScheduleModal();
    }
  });
  els.settingsCancelBtn.addEventListener('click', closeSettingsModal);
  els.settingsForm.addEventListener('submit', saveSettings);
  els.settingsModal.addEventListener('click', (event) => {
    if (event.target === els.settingsModal) {
      closeSettingsModal();
    }
  });
  els.logoutBtn.addEventListener('click', handleLogout);
  els.refreshProjectsBtn.addEventListener('click', refreshProjects);
  els.refreshFilesBtn.addEventListener('click', refreshFiles);
  els.prevProjectsBtn.addEventListener('click', () => {
    if (state.paging.loading || state.syncingAll) return;
    const nextOffset = Math.max(0, state.paging.offset - state.paging.pageSize);
    loadProjects(nextOffset);
  });
  els.nextProjectsBtn.addEventListener('click', () => {
    if (state.paging.loading || state.syncingAll || !canGoNextPage()) return;
    const nextOffset = state.paging.offset + state.paging.pageSize;
    loadProjects(nextOffset);
  });

  async function loadUploadHistoryIndex({ rebuild = false } = {}) {
    try {
      const url = rebuild
        ? '/api/projects/upload-history?rebuild=1'
        : '/api/projects/upload-history';
      const response = await apiFetch(url);
      const data = await response.json();
      if (!response.ok || !data.success) {
        throw new Error(data.error || 'Failed to load upload history');
      }
      state.uploadedProjectIds = new Set((data.uploadedProjectIds || []).map(String));
      state.uploadHistoryLoaded = true;

      for (const page of state.paging.pageCache.values()) {
        for (const project of page?.projects || []) {
          const alreadyUploaded = state.uploadedProjectIds.has(String(project.projectId));
          project.alreadyUploaded = alreadyUploaded;
          project.isNew = !alreadyUploaded;
        }
      }
      for (const project of state.projects) {
        const alreadyUploaded = state.uploadedProjectIds.has(String(project.projectId));
        project.alreadyUploaded = alreadyUploaded;
        project.isNew = !alreadyUploaded;
      }
      renderProjects();
      updateSyncAllControls();
    } catch (error) {
      console.warn('Could not load project upload history', error);
    }
  }

  updateSyncAllControls();
  updateRefreshButtons();

  if (getSessionToken()) {
    refreshResumeSnapshot();
    updateSyncAllControls();
  }

  (async () => {
    const restored = await restoreSession();
    if (restored) {
      refreshResumeSnapshot();
      updateSyncAllControls();
      await loadUploadHistoryIndex({ rebuild: true });
      loadProjects();
    }
  })();
})();
