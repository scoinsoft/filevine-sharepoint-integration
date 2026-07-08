(() => {
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
      pauseResolver: null,
    },
  };

  const $ = (id) => document.getElementById(id);

  const els = {
    error: $('error'),
    projectsStatus: $('projects-status'),
    projectsLoadingIndicator: $('projects-loading-indicator'),
    projectsLoadingText: $('projects-loading-text'),
    projectsTotal: $('projects-total'),
    projectsList: $('projects-list'),
    prevProjectsBtn: $('prev-projects-btn'),
    nextProjectsBtn: $('next-projects-btn'),
    projectsPageInfo: $('projects-page-info'),
    filesStatus: $('files-status'),
    filesList: $('files-list'),
    syncBtn: $('sync-btn'),
    syncBtnLabel: $('sync-btn-label'),
    syncBtnSpinner: $('sync-btn-spinner'),
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
    pauseSyncAllBtn: $('pause-sync-all-btn'),
    resumeSyncAllBtn: $('resume-sync-all-btn'),
    syncAllHint: $('sync-all-hint'),
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

  function setButtonLoading(spinnerEl, labelEl, loading, idleText, loadingText) {
    if (loading) {
      spinnerEl.classList.remove('hidden');
      labelEl.textContent = loadingText;
    } else {
      spinnerEl.classList.add('hidden');
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
    const busy = state.syncing || state.syncingAll;
    els.syncBtn.disabled = !state.selected || busy || state.documents.length === 0;
    setButtonLoading(
      els.syncBtnSpinner,
      els.syncBtnLabel,
      state.syncing,
      'Send files to SharePoint',
      'Sending…'
    );
  }

  function updateSyncAllControls() {
    const count = state.paging.loadedTo;
    const ready = state.paging.allLoaded && count > 0;
    const busy = state.syncing || state.syncingAll;
    const paused = state.syncingAll && state.syncAll.paused;

    els.syncAllBtn.disabled = !ready || busy;
    setButtonLoading(
      els.syncAllBtnSpinner,
      els.syncAllBtnLabel,
      state.syncingAll && !paused,
      `Send all (${count}) projects files to SharePoint`,
      `Sending all (${count}) projects…`
    );

    els.pauseSyncAllBtn.disabled = !state.syncingAll || paused;
    els.resumeSyncAllBtn.disabled = !state.syncingAll || !paused;

    if (state.syncingAll && paused) {
      els.syncAllHint.textContent = state.syncAll.currentProjectName
        ? `Paused after current project finishes (now: ${state.syncAll.currentProjectName}). Click Resume to continue.`
        : 'Paused. Click Resume to continue.';
    } else if (state.syncingAll) {
      els.syncAllHint.textContent = state.syncAll.currentProjectName
        ? `Currently syncing: ${state.syncAll.currentProjectName}`
        : 'Sending all projects…';
    } else if (!state.paging.allLoaded) {
      els.syncAllHint.textContent = `Waiting until all projects finish loading… (${count} fetched so far)`;
    } else if (count === 0) {
      els.syncAllHint.textContent = 'No projects available to sync.';
    } else {
      els.syncAllHint.textContent = `Ready to sync ${count} project(s).`;
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
      els.projectsLoadingIndicator.classList.remove('hidden');
      els.projectsLoadingIndicator.classList.add('flex');
      els.projectsLoadingText.textContent = `Loading projects... ${state.paging.loadedTo} fetched`;
    } else {
      els.projectsLoadingIndicator.classList.add('hidden');
      els.projectsLoadingIndicator.classList.remove('flex');
    }

    if (state.paging.allLoaded) {
      els.projectsTotal.classList.remove('hidden');
      els.projectsTotal.textContent = `Total projects: ${state.paging.loadedTo}`;
    } else {
      els.projectsTotal.classList.add('hidden');
      els.projectsTotal.textContent = '';
    }

    updateSyncAllButton();
  }

  function renderProjects() {
    if (state.projects.length === 0) {
      els.projectsList.innerHTML = '';
      els.projectsStatus.textContent = 'No projects on this page.';
      updateProjectsLoadIndicator();
      updateProjectsPager();
      return;
    }

    els.projectsStatus.textContent = 'Select a project';
    updateProjectsLoadIndicator();
    els.projectsList.innerHTML = state.projects
      .map((project) => {
        const selected =
          state.selected && String(state.selected.projectId) === String(project.projectId);
        return `
          <li>
            <button
              type="button"
              data-id="${escapeHtml(project.projectId)}"
              class="w-full rounded px-3 py-2 text-left text-sm ${
                selected ? 'bg-blue-600 text-white' : 'hover:bg-gray-100'
              }"
            >
              ${escapeHtml(project.projectName)}
              <span class="${selected ? 'text-blue-100' : 'text-gray-400'} text-xs">
                · ${escapeHtml(project.projectId)}
              </span>
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
  }

  function renderFiles() {
    if (!state.selected) {
      els.filesList.innerHTML = '';
      els.filesStatus.textContent = 'Click a project to see its files.';
      return;
    }

    if (state.documents.length === 0) {
      els.filesList.innerHTML = '';
      els.filesStatus.textContent = `No files in “${state.selected.projectName}”.`;
      return;
    }

    els.filesStatus.textContent = `${state.documents.length} file(s) in “${state.selected.projectName}”`;
    els.filesList.innerHTML = state.documents
      .map(
        (doc) => `
          <li class="rounded border border-gray-100 px-3 py-2">
            <div class="font-medium">${escapeHtml(doc.filename)}</div>
            ${
              doc.folderName
                ? `<div class="text-xs text-gray-400">${escapeHtml(doc.folderName)}</div>`
                : ''
            }
          </li>
        `
      )
      .join('');
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
      const response = await fetch(`/api/projects?${params.toString()}`);
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
        renderProjects();
      }
      state.paging.allLoaded = true;
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
      els.projectsStatus.textContent = 'Loading projects…';
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
      els.projectsStatus.textContent = 'Could not load projects.';
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

    els.filesStatus.textContent = `Loading files for “${project.projectName}”…`;
    els.filesList.innerHTML = '';
    showError('');

    try {
      const response = await fetch(
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
      els.filesStatus.textContent = 'Could not load files.';
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
      els.syncAllProgressText.textContent = `Paused · ${projectsDone} / ${projectsTotal} projects`;
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

    return (async function pump() {
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
            onEvent(event, JSON.parse(dataLines.join('\n')));
          } catch {
            // ignore bad chunks
          }
        }
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

    const response = await fetch(
      `/api/projects/${encodeURIComponent(project.projectId)}/sync`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projectName: project.projectName }),
      }
    );

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
    if (!state.selected || state.syncing || state.syncingAll || state.documents.length === 0) {
      return;
    }

    state.syncing = true;
    updateSyncButton();
    updateSyncAllButton();
    showError('');

    els.progressBox.classList.remove('hidden');
    els.resultList.innerHTML = '';
    updateUploadProgress(state.documents.length, 0, 0);

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
      setProgress(0, 'Error');
      showError(error.message);
    } finally {
      state.syncing = false;
      updateSyncButton();
      updateSyncAllButton();
      renderProjects();
    }
  }

  async function startSyncAll() {
    if (state.syncing || state.syncingAll || !state.paging.allLoaded) return;

    const allProjects = getAllCachedProjects();
    if (allProjects.length === 0) return;

    state.syncingAll = true;
    state.syncAll = {
      projectsTotal: allProjects.length,
      projectsDone: 0,
      filesOk: 0,
      filesFail: 0,
      currentProjectName: '',
      paused: false,
      pauseResolver: null,
    };

    updateSyncButton();
    updateSyncAllControls();
    updateProjectsPager();
    showError('');

    els.syncAllProgressBox.classList.remove('hidden');
    els.syncAllResultList.innerHTML = '';
    updateSyncAllProgressUi();
    addSyncAllHistory(`Starting sync for ${allProjects.length} project(s)…`, 'info');

    for (const project of allProjects) {
      await waitIfPaused();

      state.syncAll.currentProjectName = project.projectName || `Project ${project.projectId}`;
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
        updateSyncAllProgressUi();
        addSyncAllHistory(
          `Done: ${project.projectName} · ${summary.succeeded} uploaded, ${summary.failed} failed`,
          summary.failed > 0 ? 'error' : 'success'
        );
      } catch (error) {
        state.syncAll.projectsDone += 1;
        updateSyncAllProgressUi();
        addSyncAllHistory(
          `Failed project: ${project.projectName} — ${error.message}`,
          'error'
        );
      }
    }

    state.syncAll.currentProjectName = '';
    state.syncAll.paused = false;
    if (state.syncAll.pauseResolver) {
      state.syncAll.pauseResolver();
      state.syncAll.pauseResolver = null;
    }
    state.syncingAll = false;
    updateSyncButton();
    updateSyncAllControls();
    updateProjectsPager();
    updateSyncAllProgressUi();
    addSyncAllHistory(
      `Finished all projects · ${state.syncAll.filesOk} files uploaded, ${state.syncAll.filesFail} failed`,
      state.syncAll.filesFail > 0 ? 'error' : 'success'
    );
  }

  els.syncBtn.addEventListener('click', startSync);
  els.syncAllBtn.addEventListener('click', startSyncAll);
  els.pauseSyncAllBtn.addEventListener('click', pauseSyncAll);
  els.resumeSyncAllBtn.addEventListener('click', resumeSyncAll);
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
  loadProjects();
})();
