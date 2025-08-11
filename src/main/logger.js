// Simple logging service for main process: capture console output, keep an in-memory buffer, and optionally write to a file.
const fs = require('fs');
const path = require('path');
const { app } = require('electron');

const MAX_BUFFER = 5000; // keep last N entries in memory
let buffer = []; // { id, ts, level, message }
let nextId = 1;
let listeners = new Set();
let logDir = null;
let logFilePath = null;
let fileStream = null;
let pendingLines = [];

function tsISO(d = new Date()) {
  try {
    return d.toISOString();
  } catch {
    return new Date().toISOString();
  }
}

function ensureDir(dir) {
  try {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  } catch {}
}

function getTodayLogPath() {
  const today = new Date();
  const y = today.getFullYear();
  const m = String(today.getMonth() + 1).padStart(2, '0');
  const d = String(today.getDate()).padStart(2, '0');
  const name = `app-${y}-${m}-${d}.log`;
  return path.join(logDir, name);
}

function openStreamIfNeeded() {
  if (!logDir) return;
  if (!logFilePath) logFilePath = getTodayLogPath();
  try {
    if (!fileStream) {
      fileStream = fs.createWriteStream(logFilePath, { flags: 'a' });
      if (pendingLines.length) {
        fileStream.write(pendingLines.join(''));
        pendingLines = [];
      }
    }
  } catch {}
}

function formatLine(entry) {
  // One line per entry; message may contain newlines which we'll preserve with indentation
  const firstLine = `[${entry.ts}] [${entry.level.toUpperCase()}] ${entry.message}`;
  return firstLine.endsWith('\n') ? firstLine : firstLine + '\n';
}

function pushEntry(level, args) {
  try {
    const parts = [];
    for (const a of args) {
      if (typeof a === 'string') parts.push(a);
      else if (a instanceof Error) parts.push(`${a.message}\n${a.stack || ''}`);
      else parts.push(safeStringify(a));
    }
    const message = parts.join(' ');
    const entry = { id: nextId++, ts: tsISO(), level, message };
    buffer.push(entry);
    if (buffer.length > MAX_BUFFER) buffer = buffer.slice(-MAX_BUFFER);

    const line = formatLine(entry);
    if (fileStream) fileStream.write(line);
    else pendingLines.push(line);

    // notify listeners
    for (const fn of listeners) {
      try { fn(entry); } catch {}
    }
  } catch {}
}

function safeStringify(obj) {
  try {
    return JSON.stringify(obj, getCircularReplacer());
  } catch {
    try { return String(obj); } catch { return '[Unserializable]'; }
  }
}

function getCircularReplacer() {
  const seen = new WeakSet();
  return (key, value) => {
    if (typeof value === 'object' && value !== null) {
      if (seen.has(value)) return '[Circular]';
      seen.add(value);
    }
    return value;
  };
}

function init() {
  try {
    logDir = path.join(app.getPath('userData'), 'logs');
    ensureDir(logDir);
    openStreamIfNeeded();
  } catch {}
}

function hookConsole() {
  const original = {
    log: console.log.bind(console),
    info: console.info ? console.info.bind(console) : console.log.bind(console),
    warn: console.warn ? console.warn.bind(console) : console.log.bind(console),
    error: console.error ? console.error.bind(console) : console.log.bind(console),
    debug: console.debug ? console.debug.bind(console) : console.log.bind(console)
  };

  console.log = (...args) => { try { original.log(...args); } catch {}; pushEntry('info', args); };
  console.info = (...args) => { try { original.info(...args); } catch {}; pushEntry('info', args); };
  console.warn = (...args) => { try { original.warn(...args); } catch {}; pushEntry('warn', args); };
  console.error = (...args) => { try { original.error(...args); } catch {}; pushEntry('error', args); };
  console.debug = (...args) => { try { original.debug(...args); } catch {}; pushEntry('debug', args); };
}

function onEntry(fn) { listeners.add(fn); return () => listeners.delete(fn); }

function getLogs(opts = {}) {
  const { sinceId = null, limit = 500, level = 'all', search = '' } = opts;
  let arr = buffer;
  if (sinceId != null) arr = arr.filter(e => e.id > sinceId);
  if (level && level !== 'all') arr = arr.filter(e => e.level === level);
  if (search && String(search).trim()) {
    const s = String(search).toLowerCase();
    arr = arr.filter(e => (e.message || '').toLowerCase().includes(s));
  }
  const out = limit > 0 ? arr.slice(-limit) : arr.slice();
  return { entries: out, totalBuffered: buffer.length, lastId: buffer.length ? buffer[buffer.length - 1].id : 0 };
}

function exportAllAsString() {
  try {
    const lines = buffer.map(formatLine);
    return lines.join('');
  } catch { return ''; }
}

module.exports = {
  init,
  hookConsole,
  onEntry,
  getLogs,
  exportAllAsString,
  get logFilePath() { return logFilePath; }
};
