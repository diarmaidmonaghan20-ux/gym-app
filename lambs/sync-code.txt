// Lamb tracker sync. Paste this into Extensions > Apps Script in your Google Sheet, then Deploy > New deployment > Web app
// (Execute as: Me, Who has access: Anyone). Keep the web app address private: anyone with it can read and change your records.
const DATA = 'LambData';

function doGet() {
  return out_({ ok: true, state: load_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const inc = JSON.parse(e.postData.contents || '{}');
    const merged = merge_(load_(), inc.state || {});
    store_(merged);
    summary_(merged);
    return out_({ ok: true, state: merged });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function merge_(a, b) {
  a = a || {}; b = b || {};
  const out = {}, dead = {};
  [a.dead || {}, b.dead || {}].forEach(d => { for (const k in d) dead[k] = Math.max(dead[k] || 0, d[k]); });
  ['groups', 'products', 'log', 'deliv'].forEach(k => {
    const m = new Map();
    (a[k] || []).forEach(r => m.set(r.id, r));
    (b[k] || []).forEach(r => { const o = m.get(r.id); if (!o || (r.u || 0) >= (o.u || 0)) m.set(r.id, r); });
    out[k] = [...m.values()].filter(r => !(dead[r.id] >= (r.u || 0)));
  });
  const sa = a.setU || 0, sb = b.setU || 0;
  out.set = sb >= sa ? (b.set || a.set || {}) : a.set;
  out.setU = Math.max(sa, sb);
  out.dead = dead;
  out.updated = Date.now();
  return out;
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function load_() {
  const sh = sheet_(DATA), n = sh.getLastRow();
  if (!n) return null;
  const txt = sh.getRange(1, 1, n, 1).getValues().map(r => String(r[0]).slice(1)).join('');
  return txt ? JSON.parse(txt) : null;
}

function store_(o) {
  const sh = sheet_(DATA), s = JSON.stringify(o), parts = [];
  for (let i = 0; i < s.length; i += 40000) parts.push(['~' + s.slice(i, i + 40000)]);
  sh.clear();
  if (parts.length) sh.getRange(1, 1, parts.length, 1).setValues(parts);
}

function summary_(o) {
  const sh = sheet_('Groups'), rows = [['Group', 'Lambs', 'Average kg', 'Last weighed', 'Move on at kg', 'Meal kg/week']];
  (o.groups || []).forEach(g => rows.push([g.name || '', g.count || 0, g.avg || '', g.weighed || '', g.target || '', g.meal || '']));
  sh.clear();
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
