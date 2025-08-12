// Downloads Microsoft.Exchange.WebServices.dll from NuGet and places it into resources/ews/
// Runs on postinstall. Soft-fails if offline.

const fs = require('fs');
const path = require('path');
const https = require('https');
const os = require('os');
const AdmZip = require('adm-zip');

const NUGET_URL = 'https://globalcdn.nuget.org/packages/microsoft.exchange.webservices.2.2.0.nupkg';
const OUT_DIR = path.join(process.cwd(), 'resources', 'ews');
const OUT_DLL = path.join(OUT_DIR, 'Microsoft.Exchange.WebServices.dll');

function ensureDir(p) { try { fs.mkdirSync(p, { recursive: true }); } catch {} }

function download(url, dest) {
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(dest);
    https.get(url, (res) => {
      if (res.statusCode !== 200) {
        file.close(() => fs.unlink(dest, () => {}));
        return reject(new Error(`HTTP ${res.statusCode} for ${url}`));
      }
      res.pipe(file);
      file.on('finish', () => file.close(() => resolve(dest)));
    }).on('error', (err) => {
      file.close(() => fs.unlink(dest, () => {}));
      reject(err);
    });
  });
}

function extractDll(nupkgPath) {
  const zip = new AdmZip(nupkgPath);
  const entries = zip.getEntries();
  const prefs = [
    /^(?:lib\/|lib\\)40\/(?:Microsoft\.Exchange\.WebServices\.dll)$/i,
    /^(?:lib\/|lib\\)45\/(?:Microsoft\.Exchange\.WebServices\.dll)$/i,
    /^(?:lib\/|lib\\)[^\/\\]+\/(?:Microsoft\.Exchange\.WebServices\.dll)$/i,
  ];
  for (const rx of prefs) {
    const e = entries.find(en => rx.test(en.entryName));
    if (e) return e.getData();
  }
  return null;
}

(async () => {
  try {
    if (fs.existsSync(OUT_DLL)) {
      console.log('[EWS] DLL already present');
      return;
    }
    ensureDir(OUT_DIR);
    const tmp = path.join(os.tmpdir(), `ews_${Date.now()}.nupkg`);
    console.log('[EWS] Downloading NuGet package…');
    await download(NUGET_URL, tmp);
    console.log('[EWS] Extracting DLL…');
    const buf = extractDll(tmp);
    if (!buf) throw new Error('DLL not found in NuGet package');
    fs.writeFileSync(OUT_DLL, buf);
    console.log('[EWS] DLL installed to resources/ews');
    try { fs.unlinkSync(tmp); } catch {}
  } catch (e) {
    console.warn('[EWS] DLL setup failed:', e.message);
  }
})();
