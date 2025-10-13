/***** CONFIG *****/
const BASE_FOLDER_ID = '1JcGcwRdusWMES0RFJN8dy0gwW1KOTBcH'; // cieľový existujúci priečinok
const EXPORT_PREFIX  = 'export_';
const KEEP_LAST      = 10;                                   // ponechaj posledných X exportov (0 = nerotovať)
const TZ             = 'Europe/Bratislava';
const MAX_RUN_MS     = 5 * 60 * 1000;                        // ~5 min na jeden beh (bezpečne pod limit)

/***** PUBLIC: spusti nový export *****/
function exportBoundProjectToDrive() {
  const base = getFolderWithRetry_(BASE_FOLDER_ID);
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HH-mm');
  const exportFolder = base.createFolder(EXPORT_PREFIX + stamp);

  // načítaj obsah projektu
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(`https://script.googleapis.com/v1/projects/${scriptId}/content`, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`Apps Script API ${res.getResponseCode()}: ${res.getContentText().substring(0, 600)}`);
  }

  const data  = JSON.parse(res.getContentText());
  const files = (data.files || []).map(f => ({
    name: f.name,
    type: f.type,
    ext:  f.type === 'HTML' ? '.html' : (f.type === 'JSON' ? '.json' : '.gs'),
    src:  f.source || ''
  }));

  // init progress
  const state = {
    stamp,
    folderId: exportFolder.getId(),
    total: files.length,
    index: 0,
    created: new Date().toISOString()
  };
  PropertiesService.getUserProperties().setProperty('EXPORT_STATE', JSON.stringify(state));

  // vytvor progress súbor
  writeProgress_(exportFolder, state, 'INIT');

  // exportuj s časovým limitom
  runChunk_(files, state);

  // rotácia starých exportov
  if (KEEP_LAST > 0) cleanupOldExports_(base, EXPORT_PREFIX, KEEP_LAST);
}

/***** PUBLIC: pokračuj v rozpracovanom exporte *****/
function exportResume() {
  const state = JSON.parse(PropertiesService.getUserProperties().getProperty('EXPORT_STATE') || '{}');
  if (!state.folderId || !state.stamp) {
    SpreadsheetApp.getUi().alert('Nebola nájdená rozpracovaná relácia exportu.');
    return;
  }

  const folder = DriveApp.getFolderById(state.folderId);

  // načítaj obsah projektu znova (istota, že máme zdroj)
  const token = ScriptApp.getOAuthToken();
  const res = UrlFetchApp.fetch(`https://script.googleapis.com/v1/projects/${ScriptApp.getScriptId()}/content`, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`Apps Script API ${res.getResponseCode()}: ${res.getContentText().substring(0, 600)}`);
  }
  const data  = JSON.parse(res.getContentText());
  const files = (data.files || []).map(f => ({
    name: f.name,
    type: f.type,
    ext:  f.type === 'HTML' ? '.html' : (f.type === 'JSON' ? '.json' : '.gs'),
    src:  f.source || ''
  }));

  runChunk_(files, state, folder);
}

/***** CORE: spracuj po častiach s limitom času *****/
function runChunk_(files, state, folderOpt) {
  const start = Date.now();
  const folder = folderOpt || DriveApp.getFolderById(state.folderId);

  for (let i = state.index; i < state.total; i++) {
    const f = files[i];

    // bezpečný zápis s backoffom
    writeFileSafe_(folder, `${f.name}${f.ext}`, f.src);

    state.index = i + 1;
    if ((i + 1) % 3 === 0 || i + 1 === state.total) {
      writeProgress_(folder, state, 'IN_PROGRESS');
    }

    // ak sme blízko limitu, uložíme stav a skončíme
    if (Date.now() - start > MAX_RUN_MS && state.index < state.total) {
      PropertiesService.getUserProperties().setProperty('EXPORT_STATE', JSON.stringify(state));
      writeProgress_(folder, state, 'PAUSED_TIME_LIMIT');
      SpreadsheetApp.getUi().alert(`⏸️ Pauza po ${state.index}/${state.total} súboroch. Spusť "exportResume()" pre pokračovanie.`);
      return;
    }
  }

  // hotovo
  writeProgress_(folder, state, 'DONE');
  PropertiesService.getUserProperties().deleteProperty('EXPORT_STATE');
  SpreadsheetApp.getUi().alert(`✅ Export hotový: ${folder.getUrl()}`);
}

/***** HELPERS *****/
function writeFileSafe_(folder, name, content) {
  // ak existuje, zahoď a vytvor nanovo (jednoduché „prepísanie“)
  const it = folder.getFilesByName(name);
  while (it.hasNext()) it.next().setTrashed(true);

  // zápis s retry (Drive vie vrátiť dočasný error)
  let wait = 300;
  for (let t = 0; t < 5; t++) {
    try {
      folder.createFile(name, content, MimeType.PLAIN_TEXT);
      return;
    } catch (e) {
      const msg = String(e);
      if (!/server error|Service error|Internal|Timed out|Backend Error/i.test(msg) && t >= 2) {
        throw e;
      }
      Utilities.sleep(wait);
      wait *= 2;
    }
  }
  // posledný pokus
  folder.createFile(name, content, MimeType.PLAIN_TEXT);
}

function writeProgress_(folder, state, status) {
  const name = 'export_progress.txt';
  const body = [
    `Status: ${status}`,
    `Folder: ${state.folderId}`,
    `Stamp:  ${state.stamp}`,
    `Index:  ${state.index}/${state.total}`,
    `Time:   ${Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss')}`
  ].join('\n');

  const it = folder.getFilesByName(name);
  if (it.hasNext()) {
    const f = it.next();
    f.setTrashed(true); // jednoduchá „aktualizácia“
  }
  folder.createFile(name, body, MimeType.PLAIN_TEXT);
}

function getFolderWithRetry_(id) {
  let wait = 400;
  for (let i = 0; i < 5; i++) {
    try { return DriveApp.getFolderById(id); }
    catch (e) {
      const msg = String(e);
      if (!/server error|Service error|Internal|Timed out/i.test(msg) && i > 1) throw e;
      Utilities.sleep(wait);
      wait *= 2;
    }
  }
  return DriveApp.getFolderById(id);
}

function cleanupOldExports_(baseFolder, prefix, keepLast) {
  const list = [];
  const it = baseFolder.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    const n = f.getName();
    if (n.startsWith(prefix)) list.push({ n, f });
  }
  list.sort((a, b) => b.n.localeCompare(a.n));
  list.slice(keepLast).forEach(x => { try { x.f.setTrashed(true); } catch (_) {} });
}
