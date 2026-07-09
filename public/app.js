(() => {
  const SESSION_KEY = 'fv_sp_session_token';

  const state = {
    projects: [],
    selected: null,
    documents: [],
    syncing: false,
    syncingAll: false,
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
      filesFail: 0,
      currentProjectName: '',
      paused: false,
      pausedDueToCrash: false,
      pauseReason: '',
      pauseResolver: null,
      remainingProjects: [],
      processing: false,
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
    settingsPowerAutomateUploadUrl: $('settings-power-automate-upload-url'),
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
    statTotal: $('stat-total'),
    statOk: $('stat-ok'),
    statFail: $('stat-fail'),
    resultList: $('result-list'),
    syncAllBtn: $('sync-all-btn'),
    syncAllBtnLabel: $('sync-all-btn-label'),
    syncAllBtnSpinner: $('sync-all-btn-spinner'),
    syncAllBtnIcon: $('sync-all-btn-icon'),
    pauseSyncAllBtn: $('pause-sync-all-btn'),
    resumeSyncAllBtn: $('resume-sync-all-btn'),
    syncAllHintSpinner: $('sync-all-hint-spinner'),
    syncAllHintText: $('sync-all-hint-text'),
    syncAllProgressBox: $('sync-all-progress-box'),
    syncAllProgressText: $('sync-all-progress-text'),
    syncAllProgressPercent: $('sync-all-progress-percent'),
    syncAllProgressBar: $('sync-all-progress-bar'),
    syncAllProjectsDone: $('sync-all-projects-done'),
    syncAllProjectsTotal: $('sync-all-projects-total'),
    syncAllFilesOk: $('sync-all-files-ok'),
    syncAllFilesFail: $('sync-all-files-fail'),
    syncAllResultList: $('sync-all-result-list'),
  };

  function getSessionToken() {
    return localStorage.getItem(SESSION_KEY);
  }

  function setSessionToken(token) {
    localStorage.setItem(SESSION_KEY, token);
  }

  function clearSessionToken() {
    localStorage.removeItem(SESSION_KEY);
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
    els.settingsPowerAutomateUploadUrl.value = settings.powerAutomateUploadUrl || '';
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
      powerAutomateUploadUrl: els.settingsPowerAutomateUploadUrl.value.trim(),
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

    els.syncAllBtn.disabled = scheduledBlocked || serverSyncing || !ready || busy;
    setButtonLoading(
      els.syncAllBtnSpinner,
      els.syncAllBtnLabel,
      state.syncingAll && !paused,
      `Sync all to SharePoint (${count})`,
      `Syncing all (${count})…`,
      els.syncAllBtnIcon
    );

    els.pauseSyncAllBtn.disabled = !state.syncingAll || paused;
    els.resumeSyncAllBtn.disabled = !canResumeAll;

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
    } else if (scheduledBlocked) {
      setSyncAllHint('Scheduled sync is running — manual uploads are disabled (about 2–3 hours).', true);
    } else if (!state.paging.allLoaded) {
      setSyncAllHint(`Loading projects… (${count} loaded)`, true);
    } else if (count === 0) {
      setSyncAllHint('No projects available to sync.', false);
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

  function pauseSyncAllDueToCrash(error, remainingProjects) {
    if (!state.syncingAll) return;

    const message = getBackendPauseMessage(error);
    state.syncAll.paused = true;
    state.syncAll.pausedDueToCrash = true;
    state.syncAll.pauseReason = message;
    state.syncAll.remainingProjects = remainingProjects || [];
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
      const totalNote = state.paging.allLoaded ? ` · ${state.paging.loadedTo} total` : '';
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
        return `
          <li>
            <button
              type="button"
              data-id="${escapeHtml(project.projectId)}"
              class="flex w-full items-center gap-2 px-3 py-2.5 text-left text-sm transition ${
                selected
                  ? 'border-l-2 border-blue-600 bg-blue-50 font-medium text-blue-900'
                  : 'border-l-2 border-transparent text-gray-800 hover:bg-gray-50'
              }"
            >
              <span class="w-8 shrink-0 text-xs tabular-nums text-gray-400">${rowNumber}</span>
              <span class="shrink-0 rounded bg-gray-100 px-1.5 py-0.5 text-xs tabular-nums text-gray-600 ${
                selected ? 'bg-blue-100 text-blue-800' : ''
              }">${escapeHtml(projectNumber)}</span>
              <span class="min-w-0 flex-1 truncate">${escapeHtml(project.projectName)}</span>
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
        projects: data.projects || [],
        loadedTo: parseNonNegativeInt(
          data.loadedTo,
          normalizedOffset + (data.projects || []).length
        ),
        hasMore: Boolean(data.hasMore),
      };
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

  function addResult(filename, ok, detail) {
    const li = document.createElement('li');
    li.className = ok
      ? 'rounded bg-green-50 px-2 py-1 text-green-800'
      : 'rounded bg-red-50 px-2 py-1 text-red-800';
    li.innerHTML = ok
      ? `✓ ${escapeHtml(filename)}`
      : `✗ ${escapeHtml(filename)} — ${escapeHtml(detail || 'failed')}`;
    els.resultList.prepend(li);
  }

  function updateSyncAllProgressUi() {
    const { projectsTotal, projectsDone, filesOk, filesFail, currentProjectName, paused } =
      state.syncAll;
    const percent =
      projectsTotal === 0 ? 100 : Math.round((projectsDone / projectsTotal) * 100);

    els.syncAllProjectsDone.textContent = String(projectsDone);
    els.syncAllProjectsTotal.textContent = String(projectsTotal);
    els.syncAllFilesOk.textContent = String(filesOk);
    els.syncAllFilesFail.textContent = String(filesFail);
    els.syncAllProgressBar.style.width = `${percent}%`;
    els.syncAllProgressPercent.textContent = `${percent}%`;

    if (state.syncingAll && paused) {
      els.syncAllProgressText.textContent = state.syncAll.pausedDueToCrash
        ? `Paused (backend issue) · ${projectsDone} / ${projectsTotal} projects`
        : `Paused · ${projectsDone} / ${projectsTotal} projects`;
    } else if (state.syncingAll && currentProjectName) {
      els.syncAllProgressText.textContent = `${projectsDone} / ${projectsTotal} projects · ${currentProjectName}`;
    } else if (state.syncingAll) {
      els.syncAllProgressText.textContent = `${projectsDone} / ${projectsTotal} projects`;
    } else {
      els.syncAllProgressText.textContent = `${projectsDone} / ${projectsTotal} projects done`;
    }
  }

  function addSyncAllHistory(message, type = 'info') {
    const li = document.createElement('li');
    const styles = {
      success: 'rounded bg-green-50 px-2 py-1 text-green-800',
      error: 'rounded bg-red-50 px-2 py-1 text-red-800',
      info: 'rounded bg-gray-50 px-2 py-1 text-gray-700',
    };
    li.className = styles[type] || styles.info;
    li.textContent = message;
    els.syncAllResultList.appendChild(li);
    els.syncAllResultList.scrollTop = els.syncAllResultList.scrollHeight;
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
      if (event === 'file-success' && onFileSuccess) onFileSuccess(data);
      if (event === 'file-error' && onFileError) onFileError(data);
      if (event === 'complete') {
        summary = {
          succeeded: data.succeeded || 0,
          failed: data.failed || 0,
          total: data.total || 0,
          success: Boolean(data.success),
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
    }

    try {
      await syncProjectToSharePoint(state.selected, {
        onStarted: (data) => updateUploadProgress(data.total || 0, 0, 0),
        onFileSuccess: (data) => {
          addResult(data.filename, true);
          updateUploadProgress(data.total, data.succeeded, data.failed);
        },
        onFileError: (data) => {
          addResult(data.filename, false, data.error);
          updateUploadProgress(data.total, data.succeeded, data.failed);
        },
        onComplete: (data) => {
          updateUploadProgress(data.total || 0, data.succeeded || 0, data.failed || 0);
        },
        onError: (data) => {
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

    try {
      for (let i = 0; i < projects.length; i += 1) {
        const project = projects[i];
        await waitIfPaused();

        state.syncAll.currentProjectName = project.projectName || `Project ${project.projectId}`;
        state.syncAll.remainingProjects = projects.slice(i);
        updateSyncAllControls();
        updateSyncAllProgressUi();
        addSyncAllHistory(`▶ ${state.syncAll.currentProjectName}`, 'info');

        try {
          const summary = await syncProjectToSharePoint(project, {
            onFileSuccess: (data) => {
              state.syncAll.filesOk += 1;
              updateSyncAllProgressUi();
              addSyncAllHistory(
                `  ✓ ${data.filename} (${project.projectName})`,
                'success'
              );
            },
            onFileError: (data) => {
              state.syncAll.filesFail += 1;
              updateSyncAllProgressUi();
              addSyncAllHistory(
                `  ✗ ${data.filename} — ${data.error || 'failed'} (${project.projectName})`,
                'error'
              );
            },
          });

          state.syncAll.projectsDone += 1;
          state.syncAll.remainingProjects = projects.slice(i + 1);
          updateSyncAllProgressUi();
          addSyncAllHistory(
            `Done: ${project.projectName} · ${summary.succeeded} uploaded, ${summary.failed} failed`,
            summary.failed > 0 ? 'error' : 'success'
          );
        } catch (error) {
          if (isBackendConnectionError(error)) {
            pauseSyncAllDueToCrash(error, projects.slice(i));
            return;
          }

          state.syncAll.projectsDone += 1;
          state.syncAll.remainingProjects = projects.slice(i + 1);
          updateSyncAllProgressUi();
          addSyncAllHistory(
            `Failed project: ${project.projectName} — ${error.message}`,
            'error'
          );
        }
      }

      finishSyncAll();
    } finally {
      state.syncAll.processing = false;
    }
  }

  function finishSyncAll() {
    state.syncAll.currentProjectName = '';
    state.syncAll.paused = false;
    state.syncAll.pausedDueToCrash = false;
    state.syncAll.pauseReason = '';
    state.syncAll.remainingProjects = [];
    if (state.syncAll.pauseResolver) {
      state.syncAll.pauseResolver();
      state.syncAll.pauseResolver = null;
    }
    state.syncingAll = false;
    hidePauseNotice();
    updateSyncButton();
    updateSyncAllControls();
    updateProjectsPager();
    updateSyncAllProgressUi();
    addSyncAllHistory(
      `Finished all projects · ${state.syncAll.filesOk} files uploaded, ${state.syncAll.filesFail} failed`,
      state.syncAll.filesFail > 0 ? 'error' : 'success'
    );
  }

  async function startSyncAll() {
    if (state.syncing || (state.syncingAll && !state.syncAll.paused)) return;
    if (isUploadBlockedBySchedule()) {
      showError('Uploads are disabled while the scheduled sync is running (about 2–3 hours).');
      return;
    }
    if (!state.paging.allLoaded) return;

    const allProjects = getAllCachedProjects();
    if (allProjects.length === 0) return;

    const isResume = state.syncingAll && state.syncAll.pausedDueToCrash;
    if (isResume) {
      resumeSyncAll();
      return;
    }

    state.syncingAll = true;
    state.syncAll = {
      projectsTotal: allProjects.length,
      projectsDone: 0,
      filesOk: 0,
      filesFail: 0,
      currentProjectName: '',
      paused: false,
      pausedDueToCrash: false,
      pauseReason: '',
      pauseResolver: null,
      remainingProjects: allProjects,
      processing: false,
    };

    updateSyncButton();
    updateSyncAllControls();
    updateProjectsPager();
    showError('');
    hidePauseNotice();

    els.syncAllProgressBox.classList.remove('hidden');
    els.syncAllResultList.innerHTML = '';
    updateSyncAllProgressUi();
    addSyncAllHistory(`Starting sync for ${allProjects.length} project(s)…`, 'info');

    await continueSyncAll(allProjects);
  }

  els.syncBtn.addEventListener('click', startSync);
  els.syncAllBtn.addEventListener('click', startSyncAll);
  els.pauseSyncAllBtn.addEventListener('click', pauseSyncAll);
  els.resumeSyncAllBtn.addEventListener('click', resumeSyncAll);
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

  updateSyncAllControls();
  updateRefreshButtons();

  (async () => {
    const restored = await restoreSession();
    if (restored) {
      loadProjects();
    }
  })();
})();
