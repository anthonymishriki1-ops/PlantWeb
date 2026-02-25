/***********************
 * PlantOS — Source of Truth Build (EMR Pass + Reset Toolbox)
 * Backend: Code.gs
 ***********************/

const CFG = {
  ROOT_NAME: 'PlantOS',
  INVENTORY_SHEET: 'Plant Care Tracking + Inventory',
  LOG_SHEET: 'Plant Log',
  SETTINGS_SHEET: 'PlantOS Settings',

  HEADERS: {
    UID: 'Plant UID', // CANONICAL (8-digit)
    NAME: 'Name',
    LOCATION: 'Location',
    FOLDER_ID: 'Folder ID',
    QR_FILE_ID: 'QR File ID',
    QR_URL: 'QR URL',
    LAST_WATERED: 'Last Watered',
    LAST_FERT: 'Last Fertilized',
    LAST_REPOT: 'Last Repotted',
    NOTES: 'Notes',
    REMINDER: 'Water Reminder',
    EVERY: 'Water Every (Days)',
    DUE: 'Next Water Due',
    BIRTHDAY: 'Birthday',
    PID: 'Plant ID' // Legacy
  },

  DRIVE: {
    PLANT_SUBFOLDERS: ['Photos', 'Care Notes', 'Props', 'Receipts', 'Family Photos', 'Files'],
    ROOT_FOLDERS: { PLANTS: 'Plants', QR: 'QR Master' },
    PLANTS_FOLDERS: { LOCATIONS: 'Locations' },
    QR_FOLDERS: { SYSTEM: 'System', LOCATIONS: 'Locations', PLANTS: 'Plants' }
  },

  QR_SIZE_PX: 420
};

/** WEB APP ROUTING **/
function doGet(e) {
  const q = (e && e.parameter) || {};
  const mode = String(q.mode || 'home').trim().toLowerCase();

  const raw = String(q.uid || q.pid || q.id || '').trim();
  const loc = String(q.loc || '').trim();
  const openAdd = String(q.openAdd || '').trim();

  const baseUrl = getBaseUrl_();
  let uid = '';
  if (raw) uid = resolveAnyToUid_(raw);

  // Canonicalize URL once (if someone came in via legacy PID/name)
  if (raw && uid && uid !== raw && !q._r && baseUrl) {
    const params = [];
    params.push('uid=' + encodeURIComponent(uid));
    if (mode && mode !== 'home') params.push('mode=' + encodeURIComponent(mode));
    if (loc) params.push('loc=' + encodeURIComponent(loc));
    if (openAdd) params.push('openAdd=' + encodeURIComponent(openAdd));
    params.push('_r=1');
    const target = baseUrl + '?' + params.join('&');
    return HtmlService.createHtmlOutput(`<script>location.replace(${JSON.stringify(target)});</script>`)
      .setTitle('PlantOS');
  }

  const template = HtmlService.createTemplateFromFile('App');
  template.baseUrl = baseUrl;
  template.mode = mode;
  template.uid = uid;
  template.raw = raw;
  template.loc = loc;
  template.openAdd = openAdd;

  return template.evaluate()
    .setTitle('PlantOS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ============ PUBLIC METHODS USED BY UI ============ **/

function plantosPing() {
  return { ok: true, ts: new Date().toISOString(), baseUrl: getBaseUrl_() };
}

function plantosInit() {
  ensureSheets_();
  const tree = ensureDriveTree_();
  return { ok: true, tree };
}

function plantosGetDriveInfo() {
  ensureSheets_();
  const tree = ensureDriveTree_();
  return {
    ok: true,
    rootFolderUrl: tree.rootId ? `https://drive.google.com/drive/folders/${tree.rootId}` : '',
    plantsRootUrl: tree.plantsId ? `https://drive.google.com/drive/folders/${tree.plantsId}` : '',
    locationsRootUrl: tree.locationsId ? `https://drive.google.com/drive/folders/${tree.locationsId}` : '',
    qrRootUrl: tree.qrId ? `https://drive.google.com/drive/folders/${tree.qrId}` : '',
    qrSystemUrl: tree.qrSystemId ? `https://drive.google.com/drive/folders/${tree.qrSystemId}` : '',
    qrLocationsUrl: tree.qrLocationsId ? `https://drive.google.com/drive/folders/${tree.qrLocationsId}` : '',
    qrPlantsUrl: tree.qrPlantsId ? `https://drive.google.com/drive/folders/${tree.qrPlantsId}` : ''
  };
}

function plantosHome() {
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;
  const rows = data.slice(1);

  const today = startOfDay_(new Date());
  const in7 = new Date(today); in7.setDate(in7.getDate() + 7);

  const dueNow = [];
  const upcoming = [];
  const birthdays = [];

  for (const row of rows) {
    const uid = String(row[h[CFG.HEADERS.UID]] || '').trim();
    if (!uid) continue;

    const name = String(row[h[CFG.HEADERS.NAME]] || '').trim();
    const loc = String(row[h[CFG.HEADERS.LOCATION]] || '').trim();

    const isEnabled = row[h[CFG.HEADERS.REMINDER]] === true;
    const dueDate = toDate_(row[h[CFG.HEADERS.DUE]]);
    const bday = toDate_(row[h[CFG.HEADERS.BIRTHDAY]]);

    if (bday && bday.getMonth() === today.getMonth() && bday.getDate() === today.getDate()) {
      birthdays.push(name || uid);
    }

    if (isEnabled && dueDate) {
      const d0 = startOfDay_(dueDate);
      const item = {
        uid,
        name: name || uid,
        loc,
        due: formatYmd_(d0),
        every: row[h[CFG.HEADERS.EVERY]] || ''
      };

      if (d0.getTime() <= today.getTime()) dueNow.push(item);
      else if (d0.getTime() <= in7.getTime()) upcoming.push(item);
    }
  }

  dueNow.sort((a, b) => String(a.due).localeCompare(String(b.due)));
  upcoming.sort((a, b) => String(a.due).localeCompare(String(b.due)));

  const recent = plantosGetRecentLog(25);

  return {
    ok: true,
    birthdays,
    dueNow: dueNow.slice(0, 12),
    upcoming: upcoming.slice(0, 12),
    recent
  };
}

function plantosListLocations() {
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;
  const set = {};

  data.slice(1).forEach(r => {
    const v = String(r[h[CFG.HEADERS.LOCATION]] || '').trim();
    if (v) set[v] = true;
  });

  return Object.keys(set).sort((a, b) => a.localeCompare(b));
}

function plantosCountPlants() {
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;
  let count = 0;
  data.slice(1).forEach(r => {
    if (String(r[h[CFG.HEADERS.UID]] || '').trim()) count++;
  });
  return { ok: true, count };
}

function plantosGetPlantsByLocation(location) {
  const loc = String(location || '').trim();
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;

  const out = [];
  data.slice(1).forEach(r => {
    const rLoc = String(r[h[CFG.HEADERS.LOCATION]] || '').trim();
    if (!loc || rLoc.toLowerCase() !== loc.toLowerCase()) return;

    const uid = String(r[h[CFG.HEADERS.UID]] || '').trim();
    if (!uid) return;

    out.push({
      uid,
      name: String(r[h[CFG.HEADERS.NAME]] || '').trim() || uid,
      location: rLoc,
      due: formatYmd_(toDate_(r[h[CFG.HEADERS.DUE]])),
      lastWatered: formatYmd_(toDate_(r[h[CFG.HEADERS.LAST_WATERED]]))
    });
  });

  out.sort((a, b) => a.name.localeCompare(b.name));
  return out;
}

function plantosSearch(query, limit) {
  const q = String(query || '').trim().toLowerCase();
  const lim = Math.max(1, Math.min(50, Number(limit || 15)));
  if (!q) return [];

  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const uid = String(r[h[CFG.HEADERS.UID]] || '').trim();
    if (!uid) continue;

    const name = String(r[h[CFG.HEADERS.NAME]] || '').trim();
    const pid = String(r[h[CFG.HEADERS.PID]] || '').trim();
    const loc = String(r[h[CFG.HEADERS.LOCATION]] || '').trim();

    const blob = [uid, name, pid, loc].join(' ').toLowerCase();
    if (blob.includes(q)) {
      results.push({ uid, name: name || uid, pid, location: loc });
      if (results.length >= lim) break;
    }
  }

  return results;
}

function plantosGetPlantChart(uid) {
  const key = String(uid || '').trim();
  if (!key) return { ok: false, reason: 'missing_uid' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const row = found.row;
  const h = found.h;

  const folderId = String(row[h[CFG.HEADERS.FOLDER_ID]] || '').trim();
  const qrFileId = String(row[h[CFG.HEADERS.QR_FILE_ID]] || '').trim();
  const qrUrl = String(row[h[CFG.HEADERS.QR_URL]] || '').trim();

  return {
    ok: true,
    plant: {
      uid: String(row[h[CFG.HEADERS.UID]] || '').trim(),
      pid: String(row[h[CFG.HEADERS.PID]] || '').trim(),
      name: String(row[h[CFG.HEADERS.NAME]] || '').trim(),
      location: String(row[h[CFG.HEADERS.LOCATION]] || '').trim(),
      notes: String(row[h[CFG.HEADERS.NOTES]] || ''),
      reminder: row[h[CFG.HEADERS.REMINDER]] === true,
      every: String(row[h[CFG.HEADERS.EVERY]] || ''),
      due: formatYmd_(toDate_(row[h[CFG.HEADERS.DUE]])),
      lastWatered: formatYmd_(toDate_(row[h[CFG.HEADERS.LAST_WATERED]])),
      lastFertilized: formatYmd_(toDate_(row[h[CFG.HEADERS.LAST_FERT]])),
      lastRepotted: formatYmd_(toDate_(row[h[CFG.HEADERS.LAST_REPOT]])),
      birthday: formatYmd_(toDate_(row[h[CFG.HEADERS.BIRTHDAY]])),
      folderUrl: folderId ? `https://drive.google.com/drive/folders/${folderId}` : '',
      qrFileId,
      qrImageUrl: qrFileId ? `https://drive.google.com/uc?export=view&id=${qrFileId}` : '',
      qrUrl
    }
  };
}

function plantosGetTimeline(uid, limit) {
  const key = String(uid || '').trim();
  const lim = Math.max(1, Math.min(200, Number(limit || 30)));
  if (!key) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) return [];

  const lastRow = log.getLastRow();
  if (lastRow < 2) return [];

  const start = Math.max(2, lastRow - 800 + 1); // read a chunk, then filter
  const values = log.getRange(start, 1, lastRow - start + 1, 4).getValues().reverse();

  return values
    .filter(r => String(r[1] || '').trim() === key)
    .slice(0, lim)
    .map(r => ({
      ts: formatYmd_(toDate_(r[0])),
      uid: String(r[1] || ''),
      action: String(r[2] || ''),
      details: String(r[3] || '')
    }));
}

function plantosGetRecentLog(limit) {
  const lim = Math.max(1, Math.min(200, Number(limit || 25)));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) return [];

  const lastRow = log.getLastRow();
  if (lastRow < 2) return [];

  const start = Math.max(2, lastRow - lim + 1);
  const values = log.getRange(start, 1, lastRow - start + 1, 4).getValues().reverse();

  return values.map(r => ({
    ts: formatYmd_(toDate_(r[0])),
    uid: String(r[1] || ''),
    action: String(r[2] || ''),
    details: String(r[3] || '')
  }));
}

function plantosCreatePlant(form) {
  ensureSheets_();
  ensureDriveTree_();

  const f = normalizeForm_(form);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const invInfo = getInventory_();
    const inv = invInfo.sheet;
    const h = invInfo.h;

    const uid = generateUniqueUid_(inv, h);

    const folderData = ensurePlantDriveStructure_(uid, f.name, f.location);

    const newRow = new Array(inv.getLastColumn()).fill('');
    newRow[h[CFG.HEADERS.UID]] = uid;
    newRow[h[CFG.HEADERS.NAME]] = f.name || '';
    newRow[h[CFG.HEADERS.LOCATION]] = f.location || '';
    newRow[h[CFG.HEADERS.NOTES]] = f.notes || '';
    newRow[h[CFG.HEADERS.REMINDER]] = !!f.reminder;
    newRow[h[CFG.HEADERS.EVERY]] = f.every || '';
    newRow[h[CFG.HEADERS.BIRTHDAY]] = f.birthday || '';
    newRow[h[CFG.HEADERS.FOLDER_ID]] = folderData.folderId || '';

    if (f.reminder && f.every > 0) {
      const base = startOfDay_(new Date());
      const due = new Date(base); due.setDate(due.getDate() + Number(f.every));
      newRow[h[CFG.HEADERS.DUE]] = due;
    }

    const qr = ensurePlantQr_(uid);
    newRow[h[CFG.HEADERS.QR_URL]] = qr.url || '';
    newRow[h[CFG.HEADERS.QR_FILE_ID]] = qr.fileId || '';

    inv.appendRow(newRow);
    logAction_((/^\d{8}$/.test(uid) ? uid : 'SYSTEM'), 'CREATE', `Created plant: ${f.name || uid}`);

    return { ok: true, uid };
  } finally {
    lock.releaseLock();
  }
}

function plantosUpdatePlant(uid, patch) {
  const key = String(uid || '').trim();
  if (!key) return { ok: false, reason: 'missing_uid' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const inv = found.inv;
  const h = found.h;

  const p = patch || {};
  const updates = [];

  function setField_(header, value) {
    if (!(header in h)) return;
    const col = h[header] + 1;
    inv.getRange(found.rowIndex, col).setValue(value);
    updates.push(header);
  }

  if (p.name != null) setField_(CFG.HEADERS.NAME, String(p.name || '').trim());
  if (p.location != null) setField_(CFG.HEADERS.LOCATION, String(p.location || '').trim());
  if (p.notes != null) setField_(CFG.HEADERS.NOTES, String(p.notes || ''));
  if (p.reminder != null) setField_(CFG.HEADERS.REMINDER, !!p.reminder);
  if (p.every != null) setField_(CFG.HEADERS.EVERY, p.every === '' ? '' : Number(p.every));
  if (p.birthday != null) setField_(CFG.HEADERS.BIRTHDAY, p.birthday || '');

  if (p.recomputeDue) {
    const now = new Date();
    recomputeNextWaterDue_(key, now);
    updates.push(CFG.HEADERS.DUE);
  }

  if (updates.length) logAction_(key, 'UPDATE', `Updated: ${updates.join(', ')}`);
  return { ok: true, updated: updates };
}

function plantosLogPlantAction(uid, action, details) {
  logAction_(String(uid || '').trim(), String(action || '').trim(), String(details || ''));
  return { ok: true };
}

function plantosLogPlantRepot(uid, pot, sub) {
  const details = ['Pot: ' + (pot || '-'), 'Substrate: ' + (sub || '-')].join(' | ');
  logAction_(String(uid || '').trim(), 'REPOT', details);
  return { ok: true };
}

function plantosRepairUid(uid) {
  ensureSheets_();
  ensureDriveTree_();

  const key = String(uid || '').trim();
  if (!key) return { ok: false, reason: 'missing_uid' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const row = found.row;
  const h = found.h;

  const name = String(row[h[CFG.HEADERS.NAME]] || '').trim();
  const loc = String(row[h[CFG.HEADERS.LOCATION]] || '').trim();

  const folderData = ensurePlantDriveStructure_(key, name, loc);
  if (folderData.folderId) {
    const col = h[CFG.HEADERS.FOLDER_ID] + 1;
    found.inv.getRange(found.rowIndex, col).setValue(folderData.folderId);
  }

  const qr = ensurePlantQr_(key);
  if (qr.fileId) {
    found.inv.getRange(found.rowIndex, h[CFG.HEADERS.QR_FILE_ID] + 1).setValue(qr.fileId);
    found.inv.getRange(found.rowIndex, h[CFG.HEADERS.QR_URL] + 1).setValue(qr.url);
  }

  logAction_(key, 'REPAIR', 'Repaired Drive folder + QR');
  return { ok: true, folderId: folderData.folderId || '', qr };
}

function plantosRepairAll(limit) {
  ensureSheets_();
  ensureDriveTree_();

  const invInfo = getInventory_();
  const inv = invInfo.sheet;
  const h = invInfo.h;

  const lastRow = inv.getLastRow();
  if (lastRow < 2) return { ok: true, checked: 0, fixed: 0 };

  const lim = Math.max(1, Math.min(5000, Number(limit || 5000)));
  let checked = 0;
  let fixed = 0;

  for (let r = 2; r <= lastRow && checked < lim; r++) {
    checked++;

    const uid = String(inv.getRange(r, h[CFG.HEADERS.UID] + 1).getValue() || '').trim();
    if (!/^\d{8}$/.test(uid)) continue;

    const name = String(inv.getRange(r, h[CFG.HEADERS.NAME] + 1).getValue() || '').trim();
    const loc = String(inv.getRange(r, h[CFG.HEADERS.LOCATION] + 1).getValue() || '').trim();

    let touched = false;

    const folderId = String(inv.getRange(r, h[CFG.HEADERS.FOLDER_ID] + 1).getValue() || '').trim();
    if (!folderId) {
      const fd = ensurePlantDriveStructure_(uid, name, loc);
      if (fd.folderId) {
        inv.getRange(r, h[CFG.HEADERS.FOLDER_ID] + 1).setValue(fd.folderId);
        touched = true;
      }
    }

    const qrId = String(inv.getRange(r, h[CFG.HEADERS.QR_FILE_ID] + 1).getValue() || '').trim();
    if (!qrId) {
      const qr = ensurePlantQr_(uid);
      if (qr.fileId) {
        inv.getRange(r, h[CFG.HEADERS.QR_FILE_ID] + 1).setValue(qr.fileId);
        inv.getRange(r, h[CFG.HEADERS.QR_URL] + 1).setValue(qr.url);
        touched = true;
      }
    }

    if (touched) fixed++;
  }

  logAction_('SYSTEM', 'REPAIR_ALL', `Repair all ran. Checked=${checked}, Fixed=${fixed}`);
  return { ok: true, checked, fixed };
}

function plantosUploadPhoto(uid, filename, mimeType, dataUrl) {
  const key = String(uid || '').trim();
  if (!key) return { ok: false, reason: 'missing_uid' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const folderId = String(found.row[found.h[CFG.HEADERS.FOLDER_ID]] || '').trim();
  if (!folderId) return { ok: false, reason: 'missing_folder' };

  const plantFolder = DriveApp.getFolderById(folderId);
  const photosFolder = getOrCreateChildFolder_(plantFolder, 'Photos');

  const blob = dataUrlToBlob_(dataUrl, filename || ('photo_' + Date.now()), mimeType || 'image/jpeg');
  const file = photosFolder.createFile(blob);
  file.setName(safeFileName_(filename || ('photo_' + Date.now())));

  logAction_(key, 'PHOTO', `Uploaded photo: ${file.getName()}`);
  return { ok: true, fileId: file.getId(), url: `https://drive.google.com/file/d/${file.getId()}/view` };
}

function plantosListPhotos(uid, limit) {
  const key = String(uid || '').trim();
  const lim = Math.max(1, Math.min(50, Number(limit || 12)));
  if (!key) return [];

  const found = findPlantRowByUid_(key);
  if (!found) return [];

  const folderId = String(found.row[found.h[CFG.HEADERS.FOLDER_ID]] || '').trim();
  if (!folderId) return [];

  const plantFolder = DriveApp.getFolderById(folderId);
  const photosFolder = findChildFolder_(plantFolder, 'Photos');
  if (!photosFolder) return [];

  const files = [];
  const it = photosFolder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push({ id: f.getId(), name: f.getName(), url: `https://drive.google.com/file/d/${f.getId()}/view` });
  }
  files.reverse();
  return files.slice(0, lim);
}

/** ===== RESET TOOLBOX (this is what you were missing) ===== **/

function plantosFixActiveUrl() {
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) {}
  url = String(url || '').trim();
  if (!url) return { ok: false, reason: 'service_url_unavailable' };
  ensureSheets_();
  setSetting_('ACTIVE_WEBAPP_URL', url);
  return { ok: true, url };
}

function plantosRebuildSystemQrs() {
  ensureSheets_();
  const tree = ensureDriveTree_();
  const baseUrl = getBaseUrl_();
  if (!baseUrl) return { ok: false, reason: 'missing_base_url' };

  const qrSystem = DriveApp.getFolderById(tree.qrSystemId);
  const items = [
    { name: 'Home.png', url: baseUrl + '?mode=home' },
    { name: 'Add.png', url: baseUrl + '?mode=home&openAdd=1' },
    { name: 'Locations.png', url: baseUrl + '?mode=locations' },
    { name: 'Search.png', url: baseUrl + '?mode=search' },
    { name: 'Log.png', url: baseUrl + '?mode=log' },
    { name: 'Settings.png', url: baseUrl + '?mode=settings' }
  ];

  const out = [];
  items.forEach(it => {
    const f = upsertQrPng_(qrSystem, it.name, it.url);
    out.push({ name: it.name, fileId: f.getId(), url: it.url });
  });

  logAction_('SYSTEM', 'QR_SYSTEM_REBUILD', `System QRs rebuilt: ${out.length}`);
  return { ok: true, count: out.length, items: out };
}

function plantosRebuildLocationQrs() {
  ensureSheets_();
  const tree = ensureDriveTree_();
  const baseUrl = getBaseUrl_();
  if (!baseUrl) return { ok: false, reason: 'missing_base_url' };

  const qrLoc = DriveApp.getFolderById(tree.qrLocationsId);
  const locs = plantosListLocations();
  const out = [];

  locs.forEach(loc => {
    const url = `${baseUrl}?mode=locations&loc=${encodeURIComponent(loc)}`;
    const fname = `Loc_${safeFileName_(loc)}.png`;
    const f = upsertQrPng_(qrLoc, fname, url);
    out.push({ location: loc, fileId: f.getId(), url });
  });

  logAction_('SYSTEM', 'QR_LOC_REBUILD', `Location QRs rebuilt: ${out.length}`);
  return { ok: true, count: out.length };
}

function plantosRebuildPlantQrs(limit) {
  ensureSheets_();
  ensureDriveTree_();

  const invInfo = getInventory_();
  const inv = invInfo.sheet;
  const h = invInfo.h;
  const lastRow = inv.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const lim = Math.max(1, Math.min(5000, Number(limit || 5000)));
  let count = 0;

  for (let r = 2; r <= lastRow && count < lim; r++) {
    const uid = String(inv.getRange(r, h[CFG.HEADERS.UID] + 1).getValue() || '').trim();
    if (!/^\d{8}$/.test(uid)) continue;

    const qr = ensurePlantQr_(uid);
    if (qr.fileId) {
      inv.getRange(r, h[CFG.HEADERS.QR_FILE_ID] + 1).setValue(qr.fileId);
      inv.getRange(r, h[CFG.HEADERS.QR_URL] + 1).setValue(qr.url);
    }
    count++;
  }

  logAction_('SYSTEM', 'QR_PLANT_REBUILD', `Plant QRs rebuilt: ${count}`);
  return { ok: true, count };
}

function plantosRecomputeAllDue(limit) {
  ensureSheets_();

  const invInfo = getInventory_();
  const inv = invInfo.sheet;
  const h = invInfo.h;
  const lastRow = inv.getLastRow();
  if (lastRow < 2) return { ok: true, count: 0 };

  const lim = Math.max(1, Math.min(5000, Number(limit || 5000)));
  let count = 0;

  for (let r = 2; r <= lastRow && count < lim; r++) {
    const uid = String(inv.getRange(r, h[CFG.HEADERS.UID] + 1).getValue() || '').trim();
    if (!/^\d{8}$/.test(uid)) continue;

    const enabled = inv.getRange(r, h[CFG.HEADERS.REMINDER] + 1).getValue() === true;
    const every = Number(inv.getRange(r, h[CFG.HEADERS.EVERY] + 1).getValue() || 0);
    if (!enabled || !every || every < 1) continue;

    const last = inv.getRange(r, h[CFG.HEADERS.LAST_WATERED] + 1).getValue();
    const base = toDate_(last) || new Date();
    const d = startOfDay_(base);
    d.setDate(d.getDate() + every);

    inv.getRange(r, h[CFG.HEADERS.DUE] + 1).setValue(d);
    count++;
  }

  logAction_('SYSTEM', 'DUE_RECOMPUTE', `Due dates recomputed: ${count}`);
  return { ok: true, count };
}

function plantosSoftReset() {
  ensureSheets_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SETTINGS_SHEET);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();

  const tree = ensureDriveTree_();
  logAction_('SYSTEM', 'SOFT_RESET', 'Cleared settings and rebuilt Drive tree IDs');
  return { ok: true, tree };
}

function plantosHardReset(confirmPhrase) {
  const token = String(confirmPhrase || '').trim();
  if (token !== 'NUKE_PLANTOS') return { ok: false, reason: 'confirm_required', hint: 'Type NUKE_PLANTOS' };

  ensureSheets_();

  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(CFG.ROOT_NAME);
  if (!it.hasNext()) return { ok: true, trashed: false, reason: 'no_root_folder' };

  const f = it.next();
  f.setTrashed(true);

  logAction_('SYSTEM', 'HARD_RESET', 'Trashed PlantOS folder (sheets preserved)');
  return { ok: true, trashed: true, folderId: f.getId() };
}

/** ============ LOGGING ENGINE ============ **/

function logAction_(uid, action, details) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) throw new Error('Missing log sheet: ' + CFG.LOG_SHEET);

  log.appendRow([new Date(), uid, action, details || '']);

  const key = String(uid || '').trim();
  if (/^\d{8}$/.test(key)) {
    if (action === 'WATER') {
      const now = new Date();
      updateInventoryDate_(key, CFG.HEADERS.LAST_WATERED, now);
      recomputeNextWaterDue_(key, now);
    }
    if (action === 'FERTILIZE') updateInventoryDate_(key, CFG.HEADERS.LAST_FERT, new Date());
    if (action === 'REPOT') updateInventoryDate_(key, CFG.HEADERS.LAST_REPOT, new Date());
    if (action === 'NOTE' && details) updateInventoryText_(key, CFG.HEADERS.NOTES, details);
  }
}

/** ============ HELPERS ============ **/

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let inv = ss.getSheetByName(CFG.INVENTORY_SHEET);
  if (!inv) inv = ss.insertSheet(CFG.INVENTORY_SHEET);

  let log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) log = ss.insertSheet(CFG.LOG_SHEET);

  let settings = ss.getSheetByName(CFG.SETTINGS_SHEET);
  if (!settings) settings = ss.insertSheet(CFG.SETTINGS_SHEET);

  if (log.getLastRow() === 0) log.appendRow(['Timestamp', 'UID', 'Action', 'Details']);
  if (log.getLastRow() === 1 && log.getRange(1, 1, 1, 4).getValues()[0][0] !== 'Timestamp') {
    log.insertRowBefore(1);
    log.getRange(1, 1, 1, 4).setValues([['Timestamp', 'UID', 'Action', 'Details']]);
  }

  if (settings.getLastRow() === 0) settings.appendRow(['KEY', 'VALUE']);

  const required = Object.values(CFG.HEADERS);
  const hdrs = inv.getRange(1, 1, 1, inv.getLastColumn() || required.length).getValues()[0] || [];
  if (!hdrs.filter(Boolean).length) {
    inv.clear();
    inv.getRange(1, 1, 1, required.length).setValues([required]);
  }
}

function getInventory_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.INVENTORY_SHEET);
  if (!sheet) throw new Error('Missing inventory sheet: ' + CFG.INVENTORY_SHEET);

  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const h = getHeaderMap_(headers);

  const must = [
    CFG.HEADERS.UID,
    CFG.HEADERS.NAME,
    CFG.HEADERS.LOCATION,
    CFG.HEADERS.FOLDER_ID,
    CFG.HEADERS.QR_FILE_ID,
    CFG.HEADERS.QR_URL
  ];
  for (const m of must) {
    if (!(m in h)) throw new Error('Missing required header: ' + m);
  }

  return { sheet, h };
}

function getHeaderMap_(headers) {
  const map = {};
  headers.forEach((k, i) => {
    const key = String(k || '').trim();
    if (key) map[key] = i;
  });
  return map;
}

function getBaseUrl_() {
  try {
    const u = ScriptApp.getService().getUrl();
    if (u) return u;
  } catch (e) {}

  const v = getSetting_('ACTIVE_WEBAPP_URL');
  return v || '';
}

function getSetting_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SETTINGS_SHEET);
  if (!sh) return '';

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === key) return String(data[i][1] || '').trim();
  }
  return '';
}

function setSetting_(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.SETTINGS_SHEET);
  if (!sh) throw new Error('Missing settings sheet');

  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

function resolveAnyToUid_(raw) {
  const input = String(raw || '').trim();
  if (!input) return '';

  if (/^\d{8}$/.test(input)) return input;

  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;

  const keyLower = input.toLowerCase();

  // legacy PID exact match
  for (let i = 1; i < data.length; i++) {
    const pid = String(data[i][h[CFG.HEADERS.PID]] || '').trim();
    if (pid && pid.toLowerCase() === keyLower) return String(data[i][h[CFG.HEADERS.UID]] || '').trim();
  }

  // name exact match
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][h[CFG.HEADERS.NAME]] || '').trim();
    if (name && name.toLowerCase() === keyLower) return String(data[i][h[CFG.HEADERS.UID]] || '').trim();
  }

  // name partial match
  for (let i = 1; i < data.length; i++) {
    const name = String(data[i][h[CFG.HEADERS.NAME]] || '').trim();
    if (name && name.toLowerCase().includes(keyLower)) return String(data[i][h[CFG.HEADERS.UID]] || '').trim();
  }

  return '';
}

function findPlantRowByUid_(uid) {
  const key = String(uid || '').trim();
  if (!key) return null;

  const invInfo = getInventory_();
  const inv = invInfo.sheet;
  const data = inv.getDataRange().getValues();
  const h = invInfo.h;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][h[CFG.HEADERS.UID]] || '').trim() === key) {
      return { inv, h, rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function generateUniqueUid_(invSheet, h) {
  const uidCol = h[CFG.HEADERS.UID] + 1;
  const lastRow = invSheet.getLastRow();
  const existing = new Set();

  if (lastRow >= 2) {
    const vals = invSheet.getRange(2, uidCol, lastRow - 1, 1).getValues();
    vals.forEach(v => {
      const s = String(v[0] || '').trim();
      if (s) existing.add(s);
    });
  }

  let uid = '';
  do {
    uid = Math.floor(10000000 + Math.random() * 90000000).toString();
  } while (existing.has(uid));

  return uid;
}

function ensureDriveTree_() {
  const root = findOrCreateFolder_(DriveApp.getRootFolder(), CFG.ROOT_NAME);

  const plants = findOrCreateFolder_(root, CFG.DRIVE.ROOT_FOLDERS.PLANTS);
  const locations = findOrCreateFolder_(plants, CFG.DRIVE.PLANTS_FOLDERS.LOCATIONS);

  const qr = findOrCreateFolder_(root, CFG.DRIVE.ROOT_FOLDERS.QR);
  const qrSystem = findOrCreateFolder_(qr, CFG.DRIVE.QR_FOLDERS.SYSTEM);
  const qrLocations = findOrCreateFolder_(qr, CFG.DRIVE.QR_FOLDERS.LOCATIONS);
  const qrPlants = findOrCreateFolder_(qr, CFG.DRIVE.QR_FOLDERS.PLANTS);

  setSetting_('DRIVE_ROOT_ID', root.getId());
  setSetting_('DRIVE_PLANTS_ID', plants.getId());
  setSetting_('DRIVE_LOCATIONS_ID', locations.getId());
  setSetting_('DRIVE_QR_ID', qr.getId());
  setSetting_('DRIVE_QR_SYSTEM_ID', qrSystem.getId());
  setSetting_('DRIVE_QR_LOCATIONS_ID', qrLocations.getId());
  setSetting_('DRIVE_QR_PLANTS_ID', qrPlants.getId());

  return {
    rootId: root.getId(),
    plantsId: plants.getId(),
    locationsId: locations.getId(),
    qrId: qr.getId(),
    qrSystemId: qrSystem.getId(),
    qrLocationsId: qrLocations.getId(),
    qrPlantsId: qrPlants.getId()
  };
}

function ensurePlantDriveStructure_(uid, name, location) {
  const tree = ensureDriveTree_();
  const locationsFolder = DriveApp.getFolderById(tree.locationsId);

  const locName = String(location || '').trim() || 'Unsorted';
  const locFolder = findOrCreateFolder_(locationsFolder, locName);

  const plantFolderName = `Plant ${uid} — ${String(name || '').trim() || uid}`;
  let plantFolder = findChildFolderByUid_(locFolder, uid);
  if (!plantFolder) plantFolder = locFolder.createFolder(plantFolderName);

  CFG.DRIVE.PLANT_SUBFOLDERS.forEach(sub => findOrCreateFolder_(plantFolder, sub));
  return { folderId: plantFolder.getId() };
}

function ensurePlantQr_(uid) {
  const tree = ensureDriveTree_();
  const qrPlants = DriveApp.getFolderById(tree.qrPlantsId);

  const baseUrl = getBaseUrl_();
  const url = baseUrl ? `${baseUrl}?uid=${encodeURIComponent(uid)}` : String(uid);

  const fileName = `Plant_${uid}.png`;
  const file = upsertQrPng_(qrPlants, fileName, url);

  return { url, fileId: file ? file.getId() : '' };
}

function upsertQrPng_(folder, fileName, data) {
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const size = CFG.QR_SIZE_PX;
  const api = `https://api.qrserver.com/v1/create-qr-code/?size=${size}x${size}&data=${encodeURIComponent(data)}`;
  const resp = UrlFetchApp.fetch(api, { muteHttpExceptions: true });
  const bytes = resp.getContent();
  const blob = Utilities.newBlob(bytes, 'image/png', fileName);
  return folder.createFile(blob);
}

function findOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function findChildFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function getOrCreateChildFolder_(parent, name) {
  return findOrCreateFolder_(parent, name);
}

function findChildFolderByUid_(parent, uid) {
  const it = parent.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (String(f.getName() || '').includes(uid)) return f;
  }
  return null;
}

function updateInventoryDate_(uid, header, date) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;
  if (!(header in found.h)) return;
  const col = found.h[header] + 1;
  found.inv.getRange(found.rowIndex, col).setValue(date);
}

function updateInventoryText_(uid, header, text) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;
  if (!(header in found.h)) return;
  const col = found.h[header] + 1;
  found.inv.getRange(found.rowIndex, col).setValue(text);
}

function recomputeNextWaterDue_(uid, baseDate) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;

  const every = Number(found.row[found.h[CFG.HEADERS.EVERY]] || 0);
  const enabled = found.row[found.h[CFG.HEADERS.REMINDER]] === true;
  if (!enabled || !every || every < 1) return;

  const d = startOfDay_(new Date(baseDate || new Date()));
  d.setDate(d.getDate() + every);
  updateInventoryDate_(uid, CFG.HEADERS.DUE, d);
}

function normalizeForm_(form) {
  const f = form || {};
  const name = String(f.name || '').trim();
  const location = String(f.location || '').trim();
  const notes = String(f.notes || '');

  const reminder = String(f.reminder || '').toLowerCase() === 'true' || f.reminder === true;
  const everyNum = f.every === '' || f.every == null ? 0 : Number(f.every);
  const every = Number.isFinite(everyNum) ? Math.max(0, Math.floor(everyNum)) : 0;

  const birthday = String(f.birthday || '').trim();
  return { name, location, notes, reminder, every, birthday };
}

function toDate_(v) {
  if (!v) return null;
  const d = new Date(v);
  return Number.isNaN(d.getTime()) ? null : d;
}

function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function formatYmd_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function safeFileName_(name) {
  const n = String(name || '').trim() || ('file_' + Date.now());
  return n.replace(/[\/\\:*?"<>|]+/g, '_');
}

function dataUrlToBlob_(dataUrl, filename, mimeType) {
  const s = String(dataUrl || '');
  const m = s.match(/^data:([^;]+);base64,(.+)$/);
  if (!m) throw new Error('Bad dataUrl');
  const bytes = Utilities.base64Decode(m[2]);
  return Utilities.newBlob(bytes, mimeType || m[1], filename || ('upload_' + Date.now()));
}





/***********************
 * PlantOS — SHEETS MENU (Reset/Repair toolbox in Google Sheets)
 ***********************/

/**
 * If you ALREADY have onOpen() somewhere else:
 * - rename the old one to oldOnOpen_()
 * - keep THIS onOpen() and call oldOnOpen_() inside it.
 */
function onOpen(e) {
  try { plantosBuildSheetsMenu_(); } catch (err) {}
  // If you renamed an old onOpen to oldOnOpen_(), uncomment:
  // try { oldOnOpen_(e); } catch (err) {}
}

function plantosBuildSheetsMenu_() {
  const ui = SpreadsheetApp.getUi();
  const m = ui.createMenu('PlantOS');

  m.addItem('Open Web App', 'plantosMenu_OpenWebApp_');
  m.addSeparator();

  m.addItem('Initialize (create folders + IDs)', 'plantosMenu_Init_');
  m.addItem('Fix Base URL (current deployment)', 'plantosMenu_FixUrl_');
  m.addSeparator();

  m.addItem('Repair Missing (folders + QR for rows missing data)', 'plantosMenu_RepairMissing_');
  m.addItem('Repair UID…', 'plantosMenu_RepairUidPrompt_');
  m.addSeparator();

  m.addSubMenu(
    ui.createMenu('QR Tools')
      .addItem('Rebuild System QRs', 'plantosMenu_RebuildSystemQrs_')
      .addItem('Rebuild Location QRs', 'plantosMenu_RebuildLocationQrs_')
      .addItem('Rebuild Plant QRs (all)', 'plantosMenu_RebuildPlantQrs_')
  );

  m.addSubMenu(
    ui.createMenu('Scheduling')
      .addItem('Recompute Due Dates (all)', 'plantosMenu_RecomputeDue_')
  );

  m.addSeparator();

  m.addSubMenu(
    ui.createMenu('Reset')
      .addItem('Soft Reset (clear settings, rebuild IDs)', 'plantosMenu_SoftReset_')
      .addItem('Hard Reset… (TRASH PlantOS folder)', 'plantosMenu_HardResetPrompt_')
  );

  m.addToUi();
}

/** ===== Menu actions ===== **/

function plantosMenu_OpenWebApp_() {
  const ui = SpreadsheetApp.getUi();
  const url = (typeof getBaseUrl_ === 'function' ? getBaseUrl_() : '') || '';
  if (!url) return ui.alert('No web app URL found.\n\nTry: PlantOS → Fix Base URL');
  ui.alert('Web App URL:\n\n' + url);
}

function plantosMenu_Init_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosInit === 'function') ? plantosInit() : null;
  ui.alert(res && res.ok ? 'Initialized.' : 'Init failed. Check Logs.');
}

function plantosMenu_FixUrl_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosFixActiveUrl === 'function') ? plantosFixActiveUrl() : null;
  if (res && res.ok) ui.alert('Saved base URL:\n\n' + res.url);
  else ui.alert('Could not set base URL.\n\nMake sure you are using a DEPLOYED web app.');
}

function plantosMenu_RepairMissing_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosRepairAll === 'function') ? plantosRepairAll(5000) : null;
  if (res && res.ok) ui.alert(`Repair Missing done.\n\nChecked: ${res.checked}\nFixed: ${res.fixed}`);
  else ui.alert('Repair Missing failed. Check Logs.');
}

function plantosMenu_RepairUidPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt('Repair UID', 'Enter 8-digit Plant UID:', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const uid = String(p.getResponseText() || '').trim();
  if (!/^\d{8}$/.test(uid)) return ui.alert('Needs an 8-digit UID.');
  const res = (typeof plantosRepairUid === 'function') ? plantosRepairUid(uid) : null;
  ui.alert(res && res.ok ? 'Repaired UID: ' + uid : 'Repair failed. Check Logs.');
}

function plantosMenu_RebuildSystemQrs_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosRebuildSystemQrs === 'function') ? plantosRebuildSystemQrs() : null;
  ui.alert(res && res.ok ? `System QRs rebuilt: ${res.count}` : 'Failed. Check Logs.');
}

function plantosMenu_RebuildLocationQrs_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosRebuildLocationQrs === 'function') ? plantosRebuildLocationQrs() : null;
  ui.alert(res && res.ok ? `Location QRs rebuilt: ${res.count}` : 'Failed. Check Logs.');
}

function plantosMenu_RebuildPlantQrs_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosRebuildPlantQrs === 'function') ? plantosRebuildPlantQrs(5000) : null;
  ui.alert(res && res.ok ? `Plant QRs rebuilt: ${res.count}` : 'Failed. Check Logs.');
}

function plantosMenu_RecomputeDue_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosRecomputeAllDue === 'function') ? plantosRecomputeAllDue(5000) : null;
  ui.alert(res && res.ok ? `Due dates recomputed: ${res.count}` : 'Failed. Check Logs.');
}

function plantosMenu_SoftReset_() {
  const ui = SpreadsheetApp.getUi();
  const res = (typeof plantosSoftReset === 'function') ? plantosSoftReset() : null;
  ui.alert(res && res.ok ? 'Soft reset done.' : 'Soft reset failed. Check Logs.');
}

function plantosMenu_HardResetPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const p = ui.prompt(
    'Hard Reset (DANGER)',
    'This trashes the PlantOS folder in Drive.\n\nType NUKE_PLANTOS to confirm:',
    ui.ButtonSet.OK_CANCEL
  );
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const phrase = String(p.getResponseText() || '').trim();
  const res = (typeof plantosHardReset === 'function') ? plantosHardReset(phrase) : null;

  if (res && res.ok) {
    ui.alert(res.trashed ? 'Hard reset done. PlantOS folder trashed.' : 'No PlantOS folder found to trash.');
  } else {
    ui.alert('Blocked. You must type NUKE_PLANTOS exactly.');
  }
}
