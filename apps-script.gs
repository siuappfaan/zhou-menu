// Meal Planner sync + archive (v1.8: portions, shopping lists, absolute dates).
// Replace the old script with this, then Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy.
// Meal slots are absolute: (days since 1970-01-01) * 2 + (0 lunch, 1 dinner).
const DOW = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

function doGet(e) {
  if (e && e.parameter && e.parameter.a === 'archive') return out({ weeks: readArchive() });
  return out(readState());
}
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const inc = JSON.parse(e.postData.contents);
    if (inc.type === 'archive') { upsertArchive(inc.week); return out({ ok: true }); }
    const cur = readState();
    if (!cur || inc.updatedAt >= cur.updatedAt) { writeState(inc); return out(inc); }
    return out(cur);
  } finally { lock.releaseLock(); }
}
function tab(n) { const ss = SpreadsheetApp.getActive(); return ss.getSheetByName(n) || ss.insertSheet(n); }
function kind(k) { return k === 'home' ? 'Home' : k === 'out' ? 'Eating out' : ''; }
function readState() { const v = tab('MealState').getRange('A1').getValue(); return v ? JSON.parse(v) : null; }
function dayOf(slot) { return new Date((slot >> 1) * 86400000); }
function fmt(slot, f) { return Utilities.formatDate(dayOf(slot), 'UTC', f); }

// Current week as a readable table (meals spanning several slots repeat in each)
function writeState(s) {
  tab('MealState').getRange('A1').setValue(JSON.stringify(s));
  if (!s.blocks || !s.pv) return; // old-format state: replaced as soon as the new app syncs
  const a = s.pv[0], z = s.pv[1], rows = [[fmt(a, 'EEE d MMM') + ' – ' + fmt(z, 'EEE d MMM'), s.names[0], s.names[1]]];
  for (let i = a; i <= z; i++) {
    const f = c => {
      const b = s.blocks.filter(x => x.s <= i && i < x.s + x.n && (x.l === 'b' || x.l === (c ? 'e' : 't')))[0];
      if (!b) return '';
      const sl = (b.sl || '').split('\n').map(t => t.trim()).filter(String).join(', ');
      return b.t + (b.p ? ' x' + b.p : '') + (b.k ? ' (' + kind(b.k) + ')' : '') + (b.nt ? '\n' + b.nt : '') + (sl ? '\nShopping: ' + sl : '');
    };
    rows.push([fmt(i, 'EEE d MMM') + (i % 2 ? ' Dinner' : ' Lunch'), f(0), f(1)]);
  }
  const p = tab('Meal Plan'); p.clear(); p.getRange(1, 1, rows.length, 3).setValues(rows);
}

// Archive of all past weeks (Sat to Fri): one row per meal, never deleted
const HEAD = ['Week starting', 'Date', 'Day', 'Meal', 'Meals covered', 'Who', 'Meal name', 'Where', 'Lane', 'Notes', 'Portions', 'Shopping list'];
function archiveTab() {
  const sh = tab('Meal Archive');
  if (sh.getLastRow() === 0) { sh.appendRow(HEAD); sh.setFrozenRows(1); sh.getRange('A:B').setNumberFormat('@'); }
  HEAD.forEach((h, i) => { if (sh.getRange(1, i + 1).getValue() === '') sh.getRange(1, i + 1).setValue(h); });
  return sh;
}
function upsertArchive(w) {
  const sh = archiveTab(), last = sh.getLastRow(), tz = Session.getScriptTimeZone();
  if (last > 1) {
    const v = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let r = v.length - 1; r >= 0; r--) if (String(v[r][0]) === w.ws) sh.deleteRow(r + 2);
  }
  const rows = w.blocks.map(b => [w.ws, b.d, Utilities.formatDate(new Date(b.d + 'T12:00:00Z'), 'UTC', 'EEE'),
    b.m ? 'Dinner' : 'Lunch', b.n, b.l === 'b' ? 'Both' : w.names[b.l === 't' ? 0 : 1], b.t, kind(b.k), b.l, b.nt || '', b.p || '', b.sl || '']);
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.length).setValues(rows);
}
function readArchive() {
  const sh = archiveTab(), last = sh.getLastRow(), m = {};
  if (last < 2) return [];
  sh.getRange(2, 1, last - 1, HEAD.length).getValues().forEach(r => {
    const ws = String(r[0]); (m[ws] = m[ws] || { ws: ws, blocks: [] }).blocks.push({
      d: String(r[1]), m: r[3] === 'Dinner' ? 1 : 0, n: Number(r[4]) || 1, l: String(r[8]) || 'b',
      t: String(r[6]), k: r[7] === 'Home' ? 'home' : r[7] === 'Eating out' ? 'out' : '',
      nt: String(r[9] || ''), p: Number(r[10]) || 0, sl: String(r[11] || ''), sd: [] });
  });
  return Object.keys(m).map(k => m[k]);
}
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
