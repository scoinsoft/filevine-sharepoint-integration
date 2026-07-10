const filevineService = require('./filevine.service');
const sharepointService = require('./sharepoint.service');
const { validateSharePointEnv } = require('../config/env');
const { log, logError } = require('../utils/logger');

function isArchivedProject(project) {
  return project?.isArchived === true || project?.phaseName === 'Archived';
}

async function listAllArchivedProjects(accessToken) {
  const archived = [];
  let offset = 0;
  const limit = 1000;

  while (true) {
    const page = await filevineService.listProjectsPage(accessToken, { offset, limit });
    for (const project of page.projects) {
      if (isArchivedProject(project)) {
        archived.push(project);
      }
    }

    if (!page.hasMore || page.projects.length === 0) {
      break;
    }
    offset += page.projects.length;
  }

  return archived;
}

async function countArchivedProjects() {
  const accessToken = await filevineService.authenticate();
  const archived = await listAllArchivedProjects(accessToken);
  return {
    archivedCount: archived.length,
    projects: archived.map(({ projectId, projectName, projectNumber, phaseName, isArchived }) => ({
      projectId,
      projectName,
      projectNumber,
      phaseName,
      isArchived,
    })),
  };
}

async function removeArchivedSharePointFolders(options = {}) {
  const emit = (event, data) => {
    if (typeof options.onEvent === 'function') {
      options.onEvent(event, data);
    }
  };

  validateSharePointEnv();

  emit('status', {
    stage: 'loading',
    message: 'Loading archived projects from Filevine…',
  });

  const accessToken = await filevineService.authenticate();
  await sharepointService.verifySharePointAuth();
  const archivedProjects = await listAllArchivedProjects(accessToken);
  const total = archivedProjects.length;

  emit('started', {
    total,
    message:
      total === 0
        ? 'No archived projects found in Filevine'
        : `Found ${total} archived project(s) in Filevine`,
  });

  const results = {
    deleted: [],
    skipped: [],
    failed: [],
  };

  if (total === 0) {
    const summary = {
      success: true,
      total: 0,
      deleted: 0,
      skipped: 0,
      failed: 0,
      results,
      message: 'No archived project folders to remove',
    };
    emit('complete', summary);
    return summary;
  }

  for (let index = 0; index < archivedProjects.length; index += 1) {
    const project = archivedProjects[index];
    const folderPath = sharepointService.buildProjectFolderPath(project.projectName);
    const current = index + 1;

    emit('progress', {
      stage: 'removing',
      current,
      total,
      percent: Math.round((current / total) * 100),
      projectId: project.projectId,
      projectName: project.projectName,
      folderPath,
      message: `Checking ${project.projectName}…`,
    });

    try {
      const deleteResult = await sharepointService.deleteFolderByPath(folderPath);
      if (deleteResult.notFound) {
        results.skipped.push({
          projectId: project.projectId,
          projectName: project.projectName,
          folderPath,
          reason: 'not_found',
        });
        emit('folder-skipped', {
          projectId: project.projectId,
          projectName: project.projectName,
          folderPath,
          reason: 'not_found',
          message: `No SharePoint folder: ${project.projectName}`,
        });
      } else {
        results.deleted.push({
          projectId: project.projectId,
          projectName: project.projectName,
          folderPath,
        });
        emit('folder-deleted', {
          projectId: project.projectId,
          projectName: project.projectName,
          folderPath,
          message: `Removed: ${project.projectName}`,
        });
        log('Removed archived project folder from SharePoint', {
          projectId: project.projectId,
          projectName: project.projectName,
          folderPath,
        });
      }
    } catch (error) {
      logError('Failed to remove archived SharePoint folder', error);
      results.failed.push({
        projectId: project.projectId,
        projectName: project.projectName,
        folderPath,
        error: error.message,
      });
      emit('folder-error', {
        projectId: project.projectId,
        projectName: project.projectName,
        folderPath,
        error: error.message,
        message: `Failed: ${project.projectName} — ${error.message}`,
      });
    }
  }

  const summary = {
    success: results.failed.length === 0,
    total,
    deleted: results.deleted.length,
    skipped: results.skipped.length,
    failed: results.failed.length,
    results,
    message: `Removed ${results.deleted.length} folder(s), ${results.skipped.length} not found, ${results.failed.length} failed`,
  };

  emit('complete', summary);
  return summary;
}

module.exports = {
  isArchivedProject,
  countArchivedProjects,
  removeArchivedSharePointFolders,
};
