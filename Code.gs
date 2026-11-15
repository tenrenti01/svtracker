/**
 * SV TRACKER — Google Apps Script backend
 * DPWH Bureau of Equipment
 *
 * Connects the SV Tracker web app (GitHub Pages) to the "svtracker" Google Sheet.
 *
 * SETUP (one time, from a computer or "Desktop site" mode on a phone):
 *  1. Open the svtracker Google Sheet → Extensions → Apps Script.
 *  2. Delete the sample code, paste this whole file, click Save.
 *  3. Deploy → New deployment → type "Web app".
 *       Execute as: Me
 *       Who has access: Anyone
 *  4. Authorize, then copy the Web app URL (ends with /exec).
 *  5. In the SV Tracker app, tap the Settings icon and paste the URL.
 */

const SHEET_ID = '1jNI3D30HzVtOpDsDSXtzjQP7uGPUysqmjMC6w_UbEno';
const TZ = 'Asia/Manila';

const TAB = {
  equipment: 'Equipment',
  log: 'Location_Log',
  trips: 'Trips',
  latest: 'Latest_Location'
};

/* ---------- helpers ---------- */

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rows_(sheet) {
  const v = sheet.getDataRange().getValues();
  const head = v.shift();
  return { head: head, rows: v };
}

function iso_(d) {
  return (d instanceof Date) ? Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ss") : (d || '');
}

function mapLink_(lat, lng) {
  return '=HYPERLINK("https://www.google.com/maps?q=' + lat + ',' + lng + '","Open map")';
}

function haversineKm_(a, b) {
  const R = 6371, toR = Math.PI / 180;
  const dLat = (b[0] - a[0]) * toR, dLng = (b[1] - a[1]) * toR;
  const x = Math.sin(dLat / 2) ** 2 +
    Math.cos(a[0] * toR) * Math.cos(b[0] * toR) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(x));
}

/** Adds the hidden "Client ID" column used to avoid duplicate uploads. */
function ensureClientIdColumn_(sheet) {
  const lastCol = sheet.getLastColumn();
  const head = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = head.indexOf('Client ID');
  if (idx === -1) {
    sheet.getRange(1, lastCol + 1).setValue('Client ID')
      .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1F3A68');
    idx = lastCol;
  }
  return idx + 1; // 1-based column
}

/* ---------- GET ---------- */

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'ping';
  try {
    if (action === 'ping') return json_({ ok: true, app: 'svtracker', time: iso_(new Date()) });
    if (action === 'equipment') return json_({ ok: true, equipment: getEquipment_() });
    if (action === 'history') return json_({ ok: true, history: getHistory_(e.parameter.prop, Number(e.parameter.limit) || 25) });
    return json_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function getEquipment_() {
  const ss = ss_();
  const eq = rows_(ss.getSheetByName(TAB.equipment));
  const lt = rows_(ss.getSheetByName(TAB.latest));
  const latest = {};
  lt.rows.forEach(r => {
    latest[r[0]] = { ts: iso_(r[5]), lat: r[6], lng: r[7], mode: r[8], trip: r[9] };
  });
  return eq.rows.filter(r => r[0]).map(r => ({
    prop: r[0], code: r[1], plate: r[2], type: r[3], model: r[4], year: r[5],
    office: r[6], driver: r[7], status: r[8], last: latest[r[0]] || null
  }));
}

function getHistory_(prop, limit) {
  const lg = rows_(ss_().getSheetByName(TAB.log));
  const out = [];
  for (let i = lg.rows.length - 1; i >= 0 && out.length < limit; i--) {
    const r = lg.rows[i];
    if (prop && r[1] !== prop) continue;
    out.push({ ts: iso_(r[0]), prop: r[1], mode: r[4], trip: r[5], lat: r[6], lng: r[7],
      acc: r[8], speed: r[9], source: r[10], remarks: r[11], by: r[12] });
  }
  return out;
}

/* ---------- POST ---------- */
/* Body (sent as text/plain JSON): { action: "log", records: [ {...}, ... ] } */

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action !== 'log') return json_({ ok: false, error: 'Unknown action' });
    const saved = saveRecords_(body.records || []);
    return json_({ ok: true, saved: saved });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function saveRecords_(records) {
  const ss = ss_();
  const log = ss.getSheetByName(TAB.log);
  const idCol = ensureClientIdColumn_(log);

  // Recent client IDs, to skip anything already uploaded.
  const lastRow = log.getLastRow();
  const start = Math.max(2, lastRow - 2000);
  const known = new Set(lastRow >= 2
    ? log.getRange(start, idCol, lastRow - start + 1, 1).getValues().map(r => String(r[0]))
    : []);

  const savedIds = [];
  records.forEach(rec => {
    if (!rec || !rec.clientId) return;
    if (known.has(String(rec.clientId))) { savedIds.push(rec.clientId); return; }

    const ts = new Date(rec.ts);
    const row = [
      ts, rec.prop, rec.code, rec.plate, rec.mode, rec.tripId || '',
      Number(rec.lat), Number(rec.lng),
      rec.acc === '' || rec.acc == null ? '' : Math.round(rec.acc),
      rec.speed === '' || rec.speed == null ? '' : Math.round(rec.speed * 10) / 10,
      rec.source || '', rec.remarks || '', rec.by || '', rec.device || '',
      mapLink_(rec.lat, rec.lng)
    ];
    const r = log.getLastRow() + 1;
    log.getRange(r, 1, 1, row.length).setValues([row]);
    log.getRange(r, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    log.getRange(r, idCol).setValue(rec.clientId);
    known.add(String(rec.clientId));

    updateLatest_(ss, rec, ts);
    if (rec.tripId) updateTrip_(ss, rec, ts);
    savedIds.push(rec.clientId);
  });
  return savedIds;
}

function updateLatest_(ss, rec, ts) {
  const sh = ss.getSheetByName(TAB.latest);
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) if (data[i][0] === rec.prop) { rowIdx = i + 1; break; }
  if (rowIdx > 0 && data[rowIdx - 1][5] instanceof Date && data[rowIdx - 1][5] > ts) return; // older record
  const row = [rec.prop, rec.code, rec.plate, rec.type || '', rec.office || '', ts,
    Number(rec.lat), Number(rec.lng), rec.mode, rec.tripId || '', rec.by || '', mapLink_(rec.lat, rec.lng)];
  if (rowIdx < 0) rowIdx = sh.getLastRow() + 1;
  sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  sh.getRange(rowIdx, 6).setNumberFormat('yyyy-mm-dd hh:mm');
}

function updateTrip_(ss, rec, ts) {
  const sh = ss.getSheetByName(TAB.trips);
  const data = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i = 1; i < data.length; i++) if (data[i][0] === rec.tripId) { rowIdx = i + 1; break; }

  if (rowIdx < 0) {
    rowIdx = sh.getLastRow() + 1;
    sh.getRange(rowIdx, 1, 1, 15).setValues([[
      rec.tripId, rec.prop, rec.code, rec.plate, rec.driver || rec.by || '', rec.purpose || '',
      ts, '', Number(rec.lat), Number(rec.lng), Number(rec.lat), Number(rec.lng), 1, 0, 'Ongoing'
    ]]);
    sh.getRange(rowIdx, 7, 1, 2).setNumberFormat('yyyy-mm-dd hh:mm');
    return;
  }

  const r = data[rowIdx - 1];
  const prev = [Number(r[10]), Number(r[11])];
  const dist = Number(r[13] || 0) + haversineKm_(prev, [Number(rec.lat), Number(rec.lng)]);
  const ending = /End/.test(rec.mode);
  sh.getRange(rowIdx, 8, 1, 8).setValues([[
    ending ? ts : r[7], r[8], r[9], Number(rec.lat), Number(rec.lng),
    Number(r[12] || 0) + 1, Math.round(dist * 100) / 100, ending ? 'Completed' : 'Ongoing'
  ]]);
  sh.getRange(rowIdx, 8).setNumberFormat('yyyy-mm-dd hh:mm');
}

/** Run once from the editor to test that the script can reach the sheet. */
function testSetup() {
  Logger.log(JSON.stringify(getEquipment_().slice(0, 2), null, 2));
}
