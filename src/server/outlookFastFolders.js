const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

function pwshPath64() {
  const win = process.env.WINDIR || 'C:\\Windows';
  return path.join(win, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
}

function pwshPath32() {
  // On machines 64-bit, SysWOW64 hosts 32-bit PowerShell. If it doesn't exist, fallback to 64-bit.
  const win = process.env.WINDIR || 'C:\\Windows';
  const p32 = path.join(win, 'SysWOW64', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
  return p32;
}

function resolveScript(name) {
  const dev = path.join(process.cwd(), 'scripts', name);
  if (fs.existsSync(dev)) return dev;
  const resDir = (process.resourcesPath || process.cwd());
  const prod = path.join(resDir, 'resources', 'scripts', name);
  if (fs.existsSync(prod)) return prod;
  return dev; // fallback
}

function runPwshOnce(exePath, args, timeoutMs = 15000) {
  return new Promise((resolve, reject) => {
    const ps = spawn(exePath, ['-NoProfile', '-ExecutionPolicy', 'Bypass', ...args], { windowsHide: true });
    let out = '', err = '';
    const t = setTimeout(() => { try { ps.kill('SIGTERM'); } catch {} reject(new Error('timeout')); }, timeoutMs);
    ps.stdout.on('data', d => out += d.toString());
    ps.stderr.on('data', d => err += d.toString());
    ps.on('exit', code => { clearTimeout(t); if (code === 0) resolve(out.trim()); else reject(new Error(err || `ps exited ${code}`)); });
    ps.on('error', e => { clearTimeout(t); reject(e); });
  });
}

async function runPwshFallback(args, timeoutMs = 15000) {
  // Try 64-bit first (usual case when Outlook is 64-bit). If COM class not registered or any error, try 32-bit.
  const p64 = pwshPath64();
  const p32 = pwshPath32();
  try {
    return await runPwshOnce(p64, args, timeoutMs);
  } catch (e1) {
    // Only try 32-bit if SysWOW64 path exists
    try {
      if (fs.existsSync(p32)) {
        return await runPwshOnce(p32, args, timeoutMs);
      }
      throw e1;
    } catch (e2) {
      // Prefer the second error if 32-bit tried, else first
      throw (fs.existsSync(p32) ? e2 : e1);
    }
  }
}

async function listStores() {
  const exeArgs = ['-File', resolveScript('outlook-list-stores.ps1')];
  const raw = await runPwshFallback(exeArgs, 15000);
  try { return raw ? JSON.parse(raw) : []; } catch { return []; }
}

async function listFoldersShallow(storeId, parentEntryId = '') {
  const args = ['-File', resolveScript('outlook-list-folders.ps1'), '-StoreId', storeId];
  if (parentEntryId) args.push('-ParentEntryId', parentEntryId);
  const raw = await runPwshFallback(args, 20000);
  try {
    const data = raw ? JSON.parse(raw) : null;
    // Return the whole payload so caller can access ParentName, but keep backward compatibility
    if (data && Array.isArray(data.Folders)) {
      return { parentName: data.ParentName || null, folders: data.Folders };
    }
    // Some older scripts may return an array directly
    if (Array.isArray(data)) return { parentName: null, folders: data };
    return { parentName: null, folders: [] };
  } catch {
    return { parentName: null, folders: [] };
  }
}

module.exports = { listStores, listFoldersShallow };
