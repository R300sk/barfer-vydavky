/***** === CONFIG === *****/
const OPENAI_MODEL = 'gpt-5-codex';      // názov modelu v OpenAI Responses API
const SHEET_REVIEW = 'CODEX_Review';     // názov listu s auditom

/***** === MENU === *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Codex Tools')
    .addItem('🔍 Spustiť audit projektu', 'reviewProjectWithCodex')
    .addItem('✅ Otestovať OpenAI spojenie', 'testCodexCall')
    .addItem('🧪 Otestovať Apps Script API prístup', 'testAppsScriptApi')
    .addToUi();
}

/***** === ENTRYPOINT: audit projektu === *****/
function reviewProjectWithCodex() {
  const scriptId = ScriptApp.getScriptId();
  const files = getProjectFiles_(scriptId);        // [{name,type,source}]
  const chunks = chunkFiles_(files, 55000);        // rozdelenie podľa dĺžky
  const out = [];

  const systemPrompt =
`Si senior Apps Script inžinier. Skontroluj kód (štýl, výkon, bezpečnosť, quotas, pamäť, API volania, formátovanie).
Daj výstup v štruktúre:
1) KRÁTKY SUMÁR rizík a quick-winov
2) ZOZNAM KONKRÉTNYCH OPRÁV s odkazom na file:line (ak vieš)
3) NÁVRH PATCHU ako unified diff (alebo presné blokové zmeny)
4) PRÍPADNÉ TEST/LOG príklady
Buď stručný a konkrétny.`;

  chunks.forEach((chunk, i) => {
    const prompt = composePrompt_(chunk);
    const suggestion = callCodex_(systemPrompt, prompt);
    out.push({ part: i + 1, suggestion });
  });

  writeReviewToSheet_(out);
  SpreadsheetApp.getUi().alert(`Hotovo. Audit nájdeš v liste „${SHEET_REVIEW}“.`);
}

/***** === OPENAI (Responses API) === *****/
function callCodex_(system, user) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('Chýba Script Property OPENAI_API_KEY (vložíš v Project Settings → Script properties).');

  const url = 'https://api.openai.com/v1/responses';
  const payload = {
    model: OPENAI_MODEL,
    input: [
      { role: 'system', content: system },
      { role: 'user',   content: user }
    ],
    max_output_tokens: 1200
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code !== 200) throw new Error(`OpenAI ${code}: ${text.substring(0, 600)}`);

  const data = JSON.parse(text);

  // Responses API – bezpečné vytiahnutie textu
  if (data.output && data.output.length) {
    const joined = data.output
      .map(o => (o.content?.[0]?.text || o.text || ''))
      .filter(Boolean)
      .join('\n')
      .trim();
    if (joined) return joined;
  }
  // fallback na chat/completions formát (ak by si omylom použil iný endpoint)
  if (data.choices && data.choices[0]?.message?.content) {
    return data.choices[0].message.content;
  }
  return text;
}

/***** === APPS SCRIPT API (REST) – načítanie zdrojákov projektu ===*****/
function getProjectFiles_(scriptId) {
  const token = ScriptApp.getOAuthToken();
  const url = `https://script.googleapis.com/v1/projects/${scriptId}/content`;
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();

  if (code !== 200) {
    throw new Error(`Apps Script API ${code}: ${body.substring(0, 600)}\n
Poznámka: skontroluj, že máš povolené Apps Script API v Google Cloud Console a správne oauthScopes v appsscript.json.`);
  }

  const data = JSON.parse(body);
  if (!data.files || !data.files.length) throw new Error('API vrátilo prázdne "files".');

  return data.files
    .filter(f => ['SERVER_JS', 'HTML', 'JSON'].includes(f.type))
    .map(f => ({ name: f.name, type: f.type, source: f.source || '' }));
}

/***** === PROMPT BUILDER & CHUNKING ===*****/
function composePrompt_(filesChunk) {
  let txt = 'Projektové súbory (názov, typ, obsah):\n\n';
  filesChunk.forEach(f => {
    const fence = f.type === 'HTML' ? 'html' : (f.type === 'JSON' ? 'json' : 'javascript');
    txt += `=== FILE: ${f.name} (${f.type}) ===\n`;
    txt += '```' + fence + '\n' + (f.source || '') + '\n```\n\n';
  });
  txt += 'Prosím urob audit podľa inštrukcií.';
  return txt;
}

function chunkFiles_(files, maxChars) {
  const chunks = [];
  let buf = [], len = 0;
  files.forEach(f => {
    const add = (f.source?.length || 0) + 200; // rezerva na hlavičky
    if (len + add > maxChars && buf.length) {
      chunks.push(buf); buf = []; len = 0;
    }
    buf.push(f); len += add;
  });
  if (buf.length) chunks.push(buf);
  return chunks;
}

/***** === OUTPUT DO SHEETU ===*****/
function writeReviewToSheet_(parts) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_REVIEW);
  if (!sh) sh = ss.insertSheet(SHEET_REVIEW); else sh.clear();

  sh.getRange(1, 1, 1, 2).setValues([['Part', 'Codex návrhy']]);
  const rows = parts.map(p => [p.part, p.suggestion]);
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  sh.setColumnWidths(1, 2, 480);
  sh.setFrozenRows(1);
}

/***** === TESTY ===*****/
function testCodexCall() {
  const out = callCodex_(
    'Si stručný asistent.',
    'Odpovedz jednou vetou: Test Apps Script ↔ OpenAI Responses API prešiel.'
  );
  Logger.log(out);
  SpreadsheetApp.getUi().alert('OpenAI odpoveď je v Logs (View → Logs).');
}

function testAppsScriptApi() {
  const files = getProjectFiles_(ScriptApp.getScriptId());
  SpreadsheetApp.getUi().alert(`Načítaných súborov: ${files.length}`);
  Logger.log(files.map(f => `${f.name} (${f.type})`).join('\n'));
}
