const PARTS = 'Parts';
const LOG = 'Log';
const COLS = ['ID', 'Part Number', 'Manufacturer', 'Machine Type', 'Machine', 'Description', 'Location', 'On Hand', 'First Received', 'Last Received', 'Photo'];
const LOG_COLS = ['Time', 'Part Number', 'Change', 'On Hand After', 'How', 'By', 'Log ID', 'New Part', 'Undone'];
const F = { id: 0, partNumber: 1, manufacturer: 2, machineType: 3, machine: 4, description: 5, location: 6, qty: 7, firstReceived: 8, lastReceived: 9, thumb: 10 };
const EDITABLE = ['manufacturer', 'machineType', 'machine', 'description', 'location'];

function setup() { sheets_(); }

function sheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let p = ss.getSheetByName(PARTS);
  if (!p) {
    p = ss.insertSheet(PARTS);
    p.getRange('A:G').setNumberFormat('@');
    p.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
    p.setFrozenRows(1);
  }
  let l = ss.getSheetByName(LOG);
  if (!l) {
    l = ss.insertSheet(LOG);
    l.getRange('B:B').setNumberFormat('@');
    l.getRange('G:G').setNumberFormat('@');
    l.getRange(1, 1, 1, LOG_COLS.length).setValues([LOG_COLS]).setFontWeight('bold');
    l.setFrozenRows(1);
  }
  return { p: p, l: l };
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }
function norm_(s) { return String(s == null ? '' : s).trim().replace(/\s+/g, ' '); }
function toObj_(r) {
  return { id: String(r[0]), partNumber: String(r[1]), manufacturer: String(r[2]), machineType: String(r[3]),
    machine: String(r[4]), description: String(r[5]), location: String(r[6]), qty: Number(r[7]) || 0,
    firstReceived: iso_(r[8]), lastReceived: iso_(r[9]), thumb: String(r[10] || '') };
}
function find_(p, id) {
  const n = p.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = p.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}
function readRow_(p, r) { return p.getRange(r, 1, 1, COLS.length).getValues()[0]; }
function writeRow_(p, r, row) { p.getRange(r, 1, 1, COLS.length).setValues([row]); }

/* ---------- label reading with Gemini (free tier) ---------- */
// The key lives in Project Settings → Script Properties as GEMINI_API_KEY, never in the website code.
const LABEL_PROMPT = 'This photo shows a part, or the label or packaging of a part, being received into inventory.\n' +
  'Find the part number: the manufacturer part number, MPN, P/N, material number, item number, or catalog number. ' +
  'Do not return a storage location or bin, PO number, order number, customer number, lot or batch number, serial number, date code, weight, or quantity.\n';

function readLabel_(req) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Label reading isn’t set up yet: add GEMINI_API_KEY under Project Settings → Script Properties.');
  const img = String(req.image || '');
  if (!img || img.length > 6000000) throw new Error('The photo was missing or too large. Take it again.');
  const codes = (Array.isArray(req.codes) ? req.codes : []).map(norm_).filter(String).slice(0, 10);
  const prompt = LABEL_PROMPT +
    (codes.length ? 'Barcodes in the photo decoded to these values:\n' + codes.map(function (c) { return '- ' + c; }).join('\n') +
      '\nIf one of them is the part number, use that exact value.\n' : '') +
    'Reply with only JSON like:\n' +
    '{"partNumber":"ABC-1234","description":"short plain description of the part, or empty","manufacturer":"name or empty",' +
    '"machineType":"kind of machine it is for, only if shown on the label, else empty","machine":"machine model or name, only if shown on the label, else empty","confidence":"high"}\n' +
    'confidence is "high", "medium" or "low". If no part number is readable, use "partNumber":"".';
  const text = gemini_(key, {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: 'image/jpeg', data: img } }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  });
  let r;
  try { r = JSON.parse(text.replace(/^```(?:json)?\s*|\s*```$/g, '')); } catch (e) { throw new Error('The label read came back garbled. Scan again.'); }
  const s = function (v, n) { return norm_(v).slice(0, n); };
  return { ok: true, partNumber: s(r.partNumber, 120), description: s(r.description, 200), manufacturer: s(r.manufacturer, 80),
    machineType: s(r.machineType, 80), machine: s(r.machine, 80), confidence: s(r.confidence, 10) };
}

function gemini_(key, body) {
  // Newest Flash model first; fall back if Google has renamed it
  const models = ['gemini-flash-latest', 'gemini-2.5-flash'];
  let last = '';
  for (let i = 0; i < models.length; i++) {
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent', {
      method: 'post', contentType: 'application/json', headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(body), muteHttpExceptions: true });
    const code = res.getResponseCode();
    let j = {};
    try { j = JSON.parse(res.getContentText()); } catch (e) {}
    if (code === 200) {
      const parts = (((j.candidates || [])[0] || {}).content || {}).parts || [];
      return parts.map(function (p) { return p.text || ''; }).join('');
    }
    last = (j.error && j.error.message) || ('HTTP ' + code);
    if (code === 429) throw new Error('Too many label reads right now. Wait a minute and scan again.');
    if (code !== 404) break;
  }
  throw new Error('Label reader error: ' + last);
}

// Run this once from the editor: it asks for permission to reach Gemini and checks the key works.
function testLabelReader() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Add GEMINI_API_KEY under Project Settings → Script Properties first.');
  Logger.log('Gemini says: ' + gemini_(key, { contents: [{ parts: [{ text: 'Reply with just the word OK.' }] }] }));
}

function doGet() {
  try {
    const p = sheets_().p;
    const n = p.getLastRow() - 1;
    const rows = n > 0 ? p.getRange(2, 1, n, COLS.length).getValues() : [];
    return out_({ ok: true, parts: rows.filter(function (r) { return r[0] !== ''; }).map(toObj_) });
  } catch (e) {
    return out_({ ok: false, error: String(e.message || e) });
  }
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'Bad request.' }); }
  // Reading a label doesn't touch the sheet, so it skips the lock and never holds up other people's saves
  if (req.action === 'readLabel') {
    try { return out_(readLabel_(req)); } catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const s = sheets_(), p = s.p, l = s.l;
    const by = norm_(req.by).slice(0, 60);
    const now = new Date();
    const log = function (pn, change, after, how, isNew) {
      const id = Utilities.getUuid();
      l.appendRow([now, pn, change, after, how || '', by, id, isNew ? 'yes' : '', '']);
      return id;
    };

    if (req.action === 'move') {
      const pn = norm_(req.partNumber).slice(0, 120);
      if (!pn) throw new Error('Enter a part number.');
      const id = pn.toUpperCase();
      const delta = Math.round(Number(req.delta) || 0);
      if (!delta) throw new Error('Quantity must be at least 1.');
      let r = find_(p, id), isNew = false, row;
      if (r < 0) {
        if (delta < 0) throw new Error('This part isn’t in inventory yet. Receive it first.');
        row = [id, pn, '', '', '', '', '', 0, now, now, ''];
        r = p.getLastRow() + 1;
        isNew = true;
      } else {
        row = readRow_(p, r);
      }
      const before = Number(row[F.qty]) || 0;
      const f = req.fields || {};
      EDITABLE.forEach(function (k) { if (typeof f[k] === 'string') row[F[k]] = norm_(f[k]).slice(0, 200); });
      row[F.qty] = Math.max(0, before + delta);
      if (delta > 0) row[F.lastReceived] = now;
      if (!row[F.thumb] && req.thumb && String(req.thumb).length < 45000) row[F.thumb] = String(req.thumb);
      writeRow_(p, r, row);
      const logId = log(row[F.partNumber], row[F.qty] - before, row[F.qty], req.how, isNew);
      return out_({ ok: true, part: toObj_(row), before: before, logId: logId, isNew: isNew });
    }

    if (req.action === 'undo') {
      const n = l.getLastRow() - 1;
      const ids = n > 0 ? l.getRange(2, 7, n, 1).getValues() : [];
      let lr = -1;
      for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(req.logId)) { lr = i + 2; break; }
      if (lr < 0) throw new Error('Nothing to undo.');
      const entry = l.getRange(lr, 1, 1, LOG_COLS.length).getValues()[0];
      if (entry[8]) throw new Error('That was already undone.');
      const pn = String(entry[1]), change = Number(entry[2]) || 0;
      const r = find_(p, pn.toUpperCase());
      l.getRange(lr, 9).setValue('yes');
      if (r < 0) return out_({ ok: true });
      const row = readRow_(p, r);
      if (entry[7] === 'yes' && Number(row[F.qty]) === change) {
        p.deleteRow(r);
        log(pn, -change, 0, 'undo', false);
        return out_({ ok: true, deleted: pn.toUpperCase() });
      }
      row[F.qty] = Math.max(0, (Number(row[F.qty]) || 0) - change);
      writeRow_(p, r, row);
      log(pn, -change, row[F.qty], 'undo', false);
      return out_({ ok: true, part: toObj_(row) });
    }

    const r = find_(p, String(req.id || ''));
    if (r < 0) throw new Error('That part isn’t in the sheet anymore. Refresh the page.');
    const row = readRow_(p, r);

    if (req.action === 'update') {
      const f = req.fields || {};
      EDITABLE.forEach(function (k) { if (typeof f[k] === 'string') row[F[k]] = norm_(f[k]).slice(0, 200); });
      writeRow_(p, r, row);
      return out_({ ok: true, part: toObj_(row) });
    }
    if (req.action === 'set') {
      const before = Number(row[F.qty]) || 0;
      row[F.qty] = Math.max(0, Math.round(Number(req.qty) || 0));
      writeRow_(p, r, row);
      log(row[F.partNumber], row[F.qty] - before, row[F.qty], 'count', false);
      return out_({ ok: true, part: toObj_(row) });
    }
    if (req.action === 'delete') {
      p.deleteRow(r);
      log(row[F.partNumber], -(Number(row[F.qty]) || 0), 0, 'deleted', false);
      return out_({ ok: true, deleted: String(row[F.id]) });
    }
    throw new Error('Unknown action.');
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}
