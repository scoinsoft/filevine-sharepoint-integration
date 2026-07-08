(() => {
  const state = {
    projects: [],
    selected: null,
    documents: [],
    syncing: false,
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
    progressBox: $('progress-box'),
    progressText: $('progress-text'),
    progressPercent: $('progress-percent'),
    progressBar: $('progress-bar'),
    statTotal: $('stat-total'),
    statOk: $('stat-ok'),
    statFail: $('stat-fail'),
    resultList: $('result-list'),
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

  function updateSyncButton() {
    els.syncBtn.disabled = !state.selected || state.syncing || state.documents.length === 0;
    els.syncBtn.textContent = state.syncing ? 'Sending…' : 'Send files to SharePoint';
  }

  function updateProjectsPager() {
    const start = state.projects.length === 0 ? 0 : state.paging.offset + 1;
    const end = state.paging.offset + state.projects.length;
    els.projectsPageInfo.textContent =
      state.projects.length === 0 ? '0-0' : `${start}-${end}`;
    els.prevProjectsBtn.disabled = state.paging.loading || state.paging.offset === 0;
    els.nextProjectsBtn.disabled = state.paging.loading || !canGoNextPage();
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
        if (state.syncing) return;
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

  async function startSync() {
    if (!state.selected || state.syncing || state.documents.length === 0) return;

    state.syncing = true;
    updateSyncButton();
    showError('');

    els.progressBox.classList.remove('hidden');
    els.resultList.innerHTML = '';
    updateUploadProgress(state.documents.length, 0, 0);

    try {
      const response = await fetch(
        `/api/projects/${encodeURIComponent(state.selected.projectId)}/sync`,
        {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ projectName: state.selected.projectName }),
        }
      );

      if (!response.ok) {
        const body = await response.json().catch(() => ({}));
        throw new Error(body.error || `Sync failed (${response.status})`);
      }

      await readSseStream(response, (event, data) => {
        if (event === 'started') {
          updateUploadProgress(data.total || 0, 0, 0);
        }

        if (event === 'file-success') {
          addResult(data.filename, true);
          updateUploadProgress(data.total, data.succeeded, data.failed);
        }

        if (event === 'file-error') {
          addResult(data.filename, false, data.error);
          updateUploadProgress(data.total, data.succeeded, data.failed);
        }

        if (event === 'complete') {
          updateUploadProgress(data.total || 0, data.succeeded || 0, data.failed || 0);
        }

        if (event === 'error') {
          setProgress(0, data.error || 'Sync failed');
          showError(data.error || 'Sync failed');
        }
      });
    } catch (error) {
      setProgress(0, 'Error');
      showError(error.message);
    } finally {
      state.syncing = false;
      updateSyncButton();
      renderProjects();
    }
  }

  els.syncBtn.addEventListener('click', startSync);
  els.prevProjectsBtn.addEventListener('click', () => {
    if (state.paging.loading) return;
    const nextOffset = Math.max(0, state.paging.offset - state.paging.pageSize);
    loadProjects(nextOffset);
  });
  els.nextProjectsBtn.addEventListener('click', () => {
    if (state.paging.loading || !canGoNextPage()) return;
    const nextOffset = state.paging.offset + state.paging.pageSize;
    loadProjects(nextOffset);
  });
  loadProjects();
})();
