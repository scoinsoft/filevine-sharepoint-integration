/** @type {Map<string, { projectId: string, projectName: string, startedAt: string }>} */
const activeRuns = new Map();

function startRun(projectId, projectName) {
  activeRuns.set(String(projectId), {
    projectId: String(projectId),
    projectName,
    startedAt: new Date().toISOString(),
  });
}

function endRun(projectId) {
  activeRuns.delete(String(projectId));
}

function getActiveRuns() {
  return [...activeRuns.values()];
}

function hasActiveRuns() {
  return activeRuns.size > 0;
}

module.exports = {
  startRun,
  endRun,
  getActiveRuns,
  hasActiveRuns,
};
