// Centralized resource path resolution for dev & packaged Electron environments.
// Provides resolveResource(subDirs, fileName) returning { path, tried }.
const path = require('path');
const fs = require('fs');

function buildCandidates(subDirs, fileName) {
  const parts = Array.isArray(subDirs) ? subDirs.filter(Boolean) : [subDirs];
  const joinAll = (...base) => path.join(...base, ...parts, fileName);
  const list = [];
  const envOverride = process.env;
  // Environment variable direct override (full path)
  if (envOverride && fileName === 'ews-list-folders.ps1' && envOverride.EWS_SCRIPT_PATH) {
    list.push(envOverride.EWS_SCRIPT_PATH);
  }
  if (envOverride && fileName === 'Microsoft.Exchange.WebServices.dll' && envOverride.EWS_DLL_PATH) {
    list.push(envOverride.EWS_DLL_PATH);
  }
  // Standard packaged resources
  if (process.resourcesPath) {
    list.push(joinAll(process.resourcesPath));
    list.push(joinAll(process.resourcesPath, 'powershell'));
    // Francophone installers may create 'ressources' folders; include as fallback
    list.push(joinAll(process.resourcesPath, 'ressources'));
    list.push(joinAll(process.resourcesPath, 'ressources', 'powershell'));
    // asar unpack layout (electron-builder may place extracted assets under app.asar.unpacked)
    list.push(joinAll(process.resourcesPath, 'app.asar.unpacked'));
    list.push(joinAll(process.resourcesPath, 'app.asar.unpacked', 'resources'));
    list.push(joinAll(process.resourcesPath, 'app.asar.unpacked', 'powershell'));
    // Fallback pour configuration actuelle (doublon 'resources/resources/...')
    list.push(joinAll(process.resourcesPath, 'resources'));
    list.push(joinAll(process.resourcesPath, 'resources', 'powershell'));
    list.push(joinAll(process.resourcesPath, 'ressources', 'resources'));
    list.push(joinAll(process.resourcesPath, 'ressources', 'resources', 'powershell'));
  }
  // execPath related
  try {
    const execDir = path.dirname(process.execPath || '');
    list.push(joinAll(execDir, 'resources'));
    list.push(joinAll(execDir, 'resources', 'powershell'));
    list.push(joinAll(execDir, 'ressources'));
    list.push(joinAll(execDir, 'ressources', 'powershell'));
    list.push(joinAll(execDir, 'app.asar.unpacked'));
    list.push(joinAll(execDir, 'app.asar.unpacked', 'resources'));
    list.push(joinAll(execDir, 'app.asar.unpacked', 'powershell'));
    list.push(joinAll(execDir, 'resources', 'resources'));
    list.push(joinAll(execDir, 'ressources', 'resources'));
    list.push(joinAll(execDir, 'ressources', 'resources', 'powershell'));
  } catch {}
  // app path (dev)
  try {
    const appPath = (require('electron').app?.getAppPath && require('electron').app.getAppPath()) || '';
    if (appPath) {
      list.push(joinAll(appPath));
      list.push(joinAll(appPath, 'powershell'));
      list.push(joinAll(appPath, 'resources'));
      list.push(joinAll(appPath, 'resources', 'powershell'));
    }
  } catch {}
  // cwd variants
  list.push(joinAll(process.cwd(), 'resources'));
  list.push(joinAll(process.cwd(), 'resources', 'powershell'));
  list.push(joinAll(process.cwd(), 'powershell'));
  list.push(joinAll(process.cwd()));
  return [...new Set(list)];
}

function resolveResource(subDirs, fileName) {
  const tried = buildCandidates(subDirs, fileName);
  for (const candidate of tried) {
    try { if (candidate && fs.existsSync(candidate)) return { path: candidate, tried }; } catch {}
  }
  return { path: null, tried };
}

module.exports = { resolveResource };
