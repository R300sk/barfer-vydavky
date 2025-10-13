/** ====== INDEX EXPORTU – prehľad súborov, riadkov, funkcií, triggerov ====== */
function buildExportIndex() {
  const base = DriveApp.getFolderById(BASE_FOLDER_ID);
  const latest = getLatestExportFolder_(base, EXPORT_PREFIX);
  if (!latest) throw new Error('Nenašiel som žiadny export_* podpriečinok.');

  const rows = [['Name','Type','Size (bytes)','Lines','Functions','Has onOpen/onEdit','Has ScriptApp.newTrigger','Last modified']];
  const it = latest.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!/\.(gs|html|json)$/i.test(f.getName())) continue;
    const text = f.getBlob().getDataAsString() || '';
    const type = f.getName().split('.').pop().toLowerCase();
    const lines = text ? text.split(/\r?\n/).length : 0;
    const fnMatches = [...text.matchAll(/\bfunction\s+([A-Za-z0-9_]+)\s*\(/g)].map(m => m[1]);
    const hasUITriggers = /(?:\bonOpen\b|\bonEdit\b)/.test(text);
    const hasInstallable = /ScriptApp\.newTrigger\(/.test(text);
    rows.push([
      f.getName(),
      type,
      f.getSize(),
      lines,
      fnMatches.join(', '),
      hasUITriggers ? 'yes' : '',
      hasInstallable ? 'yes' : '',
      Utilities.formatDate(f.getLastUpdated(), TZ, 'yyyy-MM-dd HH:mm')
    ]);
  }

  // vytvor Google Sheet s indexom
  const ss = SpreadsheetApp.create(`INDEX_${latest.getName()}`);
  const sh = ss.getSheets()[0];
  sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, rows[0].length);

  // ulož CSV aj do export priečinka
  const csv = rows.map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
  latest.createFile(`files_index_${latest.getName()}.csv`, csv, MimeType.CSV);

  SpreadsheetApp.getUi().alert(
    `✅ Index hotový.\nSheet: ${ss.getUrl()}\nExport priečinok: ${latest.getUrl()}`
  );
}

function getLatestExportFolder_(base, prefix) {
  const list = [];
  const it = base.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    const n = f.getName();
    if (n.startsWith(prefix)) list.push({ n, f });
  }
  if (!list.length) return null;
  list.sort((a,b) => b.n.localeCompare(a.n)); // vďaka formátu yyyy-MM-dd_HH-mm
  return list[0].f;
}
