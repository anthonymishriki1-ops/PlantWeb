/***********************
 * PlantOS — Source of Truth Build (EMR Pass + Human Names + Care Doc)
 * Backend: Code.gs
 *
 * What this fixes:
 * - QR + direct links open PLANT PAGE (no scan screen).
 * - "Plant not found" fixed (HTML calls plantosGetPlant; backend provides it).
 * - Location list never shows numeric IDs as the big label.
 * - If no nickname, location list shows taxonomy label (Genus Species + #2 if duplicates),
 *   plus DOB + Pot/Substrate line.
 * - "Add a nickname?" works and updates sheet immediately.
 * - Logging writes to Plant Log sheet + appends to Care Log Google Doc in plant folder/Care Notes.
 * - Restores Home: watering dashboard, quick log, add specimen (basic/advanced), rolling ticker.
 ***********************/

const CFG = {
  ROOT_NAME: 'PlantOS',
  INVENTORY_SHEET: 'Plant Care Tracking + Inventory',
  LOG_SHEET: 'Plant Log',
  SETTINGS_SHEET: 'PlantOS Settings',

  // IMPORTANT: header keys are TRIMMED by getHeaderMap_(), so "Nick-name  " becomes "Nick-name".
  HEADERS: {
    DISPLAY_NAME: 'Display Name',
    NICKNAME: 'Nick-name',
    NAME: 'Name',

    PLANT_ID: 'Plant ID',   // your human ID column (may be numeric in old data)
    UID: 'Plant UID',       // canonical internal 8-digit key

    CLASSIFICATION: 'Classification',
    GENUS: 'Genus',
    SPECIES: 'Species',

    LOCATION: 'Location',

    SUBSTRATE: 'Substrate',
    MEDIUM: 'Medium',
    POT_SIZE: 'Pot Size',

    BIRTHDAY: 'Birthday',

    LAST_WATERED: 'Last Watered',
    LAST_FERT: 'Last Fertilized',
    LAST_REPOT: 'Last Repotted',

    NOTES: 'Notes',

    REMINDER_ENABLED: 'Water Reminder Enabled',
    EVERY: 'Water Every (Days)',
    DUE: 'Next Water Due',

    FOLDER_ID: 'Folder ID',
    FOLDER_URL: 'Folder URL',

    QR_FILE_ID: 'QR File ID',
    QR_URL: 'QR URL',

    CARE_DOC_ID: 'Care Doc ID',
    CARE_DOC_URL: 'Care Doc URL'
  },

  DRIVE: {
    PLANT_SUBFOLDERS: ['Photos', 'Care Notes', 'Props', 'Receipts', 'Family Photos', 'Files', 'Old QR'],
    ROOT_FOLDERS: { PLANTS: 'Plants', QR: 'QR Master' },
    PLANTS_FOLDERS: { LOCATIONS: 'Locations' },
    QR_FOLDERS: { SYSTEM: 'System', LOCATIONS: 'Locations', PLANTS: 'Plants' }
  },

  QR_SIZE_PX: 420,

  // Batch safety
  MAX_BATCH_WATER: 120,
  MAX_BATCH_REPAIR: 400
};

/** ===================== WEB APP ROUTING ===================== **/
function doGet(e) {
  const q = (e && e.parameter) || {};
  let mode = String(q.mode || '').trim().toLowerCase();

  const raw = String(q.uid || q.pid || q.id || '').trim();
  const loc = String(q.loc || '').trim();
  const openAdd = String(q.openAdd || '').trim();

  const baseUrl = getBaseUrl_();
  let uid = '';
  if (raw) uid = resolveAnyToUid_(raw);

  // If someone comes via QR or direct url with ?uid=######## and no mode, go to plant page.
  if (uid && !mode) mode = 'plant';
  if (!mode) mode = 'home';

  // Canonicalize URL once if input wasn't already uid
  if (raw && uid && uid !== raw && !q._r && baseUrl) {
    const params = [];
    params.push('uid=' + encodeURIComponent(uid));
    params.push('mode=' + encodeURIComponent(mode));
    if (loc) params.push('loc=' + encodeURIComponent(loc));
    if (openAdd) params.push('openAdd=' + encodeURIComponent(openAdd));
    params.push('_r=1');
    const target = baseUrl + '?' + params.join('&');
    return HtmlService.createHtmlOutput(`<script>location.replace(${JSON.stringify(target)});</script>`).setTitle('PlantOS');
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

/** ===================== PUBLIC API FOR UI ===================== **/

function plantosPing() {
  return { ok: true, ts: new Date().toISOString(), baseUrl: getBaseUrl_() };
}

function plantosInit() {
  ensureSheets_();
  const tree = ensureDriveTree_();
  return { ok: true, tree };
}

function plantosCountPlants() {
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const uid = normalizeUid_(data[i][h[CFG.HEADERS.UID]]);
    if (uid) count++;
  }
  return { ok: true, count };
}

function plantosHome() {
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;

  const today = startOfDay_(new Date());
  const in7 = new Date(today); in7.setDate(in7.getDate() + 7);

  const dueNow = [];
  const upcoming = [];
  const birthdays = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const uid = normalizeUid_(row[h[CFG.HEADERS.UID]]);
    if (!uid) continue;

    const label = computeDisplayLabel_(row, h, /*fallbackTax=*/true);
    const enabled = row[h[CFG.HEADERS.REMINDER_ENABLED]] === true;
    const dueDate = toDate_(row[h[CFG.HEADERS.DUE]]);
    const bday = toDate_(row[h[CFG.HEADERS.BIRTHDAY]]);

    if (bday && bday.getMonth() === today.getMonth() && bday.getDate() === today.getDate()) {
      birthdays.push(label.primary);
    }

    if (enabled && dueDate) {
      const d0 = startOfDay_(dueDate);
      const item = {
        uid,
        primary: label.primary,
        secondary: label.secondary,
        due: formatYmd_(d0),
        every: String(row[h[CFG.HEADERS.EVERY]] || '').trim()
      };
      if (d0.getTime() <= today.getTime()) dueNow.push(item);
      else if (d0.getTime() <= in7.getTime()) upcoming.push(item);
    }
  }

  dueNow.sort((a, b) => String(a.due).localeCompare(String(b.due)));
  upcoming.sort((a, b) => String(a.due).localeCompare(String(b.due)));

  const recent = plantosGetRecentLog(20);

  return {
    ok: true,
    birthdays,
    dueNow: dueNow.slice(0, 20),
    upcoming: upcoming.slice(0, 20),
    recent
  };
}

function plantosListLocations() {
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;
  const set = {};
  for (let i = 1; i < data.length; i++) {
    const v = String(data[i][h[CFG.HEADERS.LOCATION]] || '').trim();
    if (v) set[v] = true;
  }
  return Object.keys(set).sort((a, b) => a.localeCompare(b));
}

/**
 * Location cards:
 * - Title line: nickname if exists; else taxonomy label (Genus Species + #2 if dup)
 * - Sub line: Genus Species (if nickname exists), otherwise Classification (if exists)
 * - Meta line: DOB + Pot/Substrate
 * - Only items with missing nickname show needsNickname=true
 */
function plantosGetPlantsByLocation(location) {
  const loc = String(location || '').trim();
  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;

  // First collect rows
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rLoc = String(row[h[CFG.HEADERS.LOCATION]] || '').trim();
    if (loc && rLoc.toLowerCase() !== loc.toLowerCase()) continue;

    const uid = normalizeUid_(row[h[CFG.HEADERS.UID]]);
    if (!uid) continue;

    rows.push({ row, uid, rLoc });
  }

  // Build duplicate counters for "no nickname" items based on taxon key
  const counts = {};
  const ordered = {}; // will hold incrementing index per key

  for (const it of rows) {
    const row = it.row;
    const nn = String(row[h[CFG.HEADERS.NICKNAME]] || '').trim();
    const dn = String(row[h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
    const nm = String(row[h[CFG.HEADERS.NAME]] || '').trim();
    const hasHumanName = !!(nn || dn || nm);
    if (hasHumanName) continue;

    const key = taxonKey_(row, h);
    counts[key] = (counts[key] || 0) + 1;
  }

  const out = [];
  for (const it of rows) {
    const row = it.row;
    const uid = it.uid;

    const nickname = String(row[h[CFG.HEADERS.NICKNAME]] || '').trim();
    const displayName = String(row[h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
    const name = String(row[h[CFG.HEADERS.NAME]] || '').trim();

    const classification = String(row[h[CFG.HEADERS.CLASSIFICATION]] || '').trim();
    const genus = String(row[h[CFG.HEADERS.GENUS]] || '').trim();
    const species = String(row[h[CFG.HEADERS.SPECIES]] || '').trim();

    const potSize = String(row[h[CFG.HEADERS.POT_SIZE]] || '').trim();
    const substrate = String(row[h[CFG.HEADERS.SUBSTRATE]] || '').trim() || String(row[h[CFG.HEADERS.MEDIUM]] || '').trim();
    const birthday = formatYmd_(toDate_(row[h[CFG.HEADERS.BIRTHDAY]]));

    const due = formatYmd_(toDate_(row[h[CFG.HEADERS.DUE]]));
    const enabled = row[h[CFG.HEADERS.REMINDER_ENABLED]] === true;

    let primary = nickname || displayName || name;

    // If no human name at all, use taxonomy and add #N if duplicates
    let suffix = '';
    if (!primary) {
      const key = taxonKey_(row, h);
      if (counts[key] > 1) {
        ordered[key] = (ordered[key] || 0) + 1;
        suffix = ' #' + ordered[key];
      }
      primary = taxonLabel_(classification, genus, species) + suffix;
    }

    // Secondary line rules:
    // - If nickname exists: show Genus Species (not repeated genus twice)
    // - If no nickname: show Classification (if exists) else Genus Species
    let secondary = '';
    const gs = [genus, species].filter(Boolean).join(' ').trim();
    if (nickname) secondary = gs || classification || '';
    else secondary = classification || gs || '';

    // Meta line: DOB + Pot/Substrate
    const metaBits = [];
    if (birthday) metaBits.push('🎂 ' + birthday);
    if (potSize) metaBits.push('🪴 ' + potSize);
    if (substrate) metaBits.push('🧱 ' + substrate);
    const meta = metaBits.join(' • ');

    out.push({
      uid,
      location: it.rLoc,
      nickname,
      displayName,
      name,
      primary,
      secondary,
      meta,
      classification,
      genus,
      species,
      potSize,
      substrate,
      birthday,
      reminderEnabled: enabled,
      due,
      needsNickname: !nickname
    });
  }

  out.sort((a, b) => String(a.primary || a.uid).localeCompare(String(b.primary || b.uid)));
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
    const row = data[i];
    const uid = normalizeUid_(row[h[CFG.HEADERS.UID]]);
    if (!uid) continue;

    const nickname = String(row[h[CFG.HEADERS.NICKNAME]] || '').trim();
    const displayName = String(row[h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
    const name = String(row[h[CFG.HEADERS.NAME]] || '').trim();
    const genus = String(row[h[CFG.HEADERS.GENUS]] || '').trim();
    const species = String(row[h[CFG.HEADERS.SPECIES]] || '').trim();
    const loc = String(row[h[CFG.HEADERS.LOCATION]] || '').trim();
    const plantId = String(row[h[CFG.HEADERS.PLANT_ID]] || '').trim();

    const primary = nickname || displayName || name || [genus, species].filter(Boolean).join(' ').trim() || 'Plant';
    const blob = [uid, plantId, nickname, displayName, name, genus, species, loc].join(' ').toLowerCase();

    if (blob.includes(q)) {
      results.push({ uid, primary, location: loc, genus, species, nickname });
      if (results.length >= lim) break;
    }
  }
  return results;
}

/** Alias to satisfy UI calls */
function plantosGetPlant(uid) {
  return plantosGetPlantChart(uid);
}

function plantosGetPlantChart(uid) {
  const key = normalizeUid_(uid);
  if (!key) return { ok: false, reason: 'missing_or_bad_uid' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const row = found.row;
  const h = found.h;

  const nickname = String(row[h[CFG.HEADERS.NICKNAME]] || '').trim();
  const displayName = String(row[h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
  const name = String(row[h[CFG.HEADERS.NAME]] || '').trim();

  const classification = String(row[h[CFG.HEADERS.CLASSIFICATION]] || '').trim();
  const genus = String(row[h[CFG.HEADERS.GENUS]] || '').trim();
  const species = String(row[h[CFG.HEADERS.SPECIES]] || '').trim();

  const plantId = String(row[h[CFG.HEADERS.PLANT_ID]] || '').trim();

  const folderId = String(row[h[CFG.HEADERS.FOLDER_ID]] || '').trim();
  const qrFileId = String(row[h[CFG.HEADERS.QR_FILE_ID]] || '').trim();
  const qrUrl = String(row[h[CFG.HEADERS.QR_URL]] || '').trim();

  const careDocId = String(row[h[CFG.HEADERS.CARE_DOC_ID]] || '').trim();
  const careDocUrl = String(row[h[CFG.HEADERS.CARE_DOC_URL]] || '').trim();

  const substrate = String(row[h[CFG.HEADERS.SUBSTRATE]] || '').trim();
  const medium = String(row[h[CFG.HEADERS.MEDIUM]] || '').trim();
  const potSize = String(row[h[CFG.HEADERS.POT_SIZE]] || '').trim();

  const reminderEnabled = row[h[CFG.HEADERS.REMINDER_ENABLED]] === true;
  const everyDays = String(row[h[CFG.HEADERS.EVERY]] || '').trim();
  const due = formatYmd_(toDate_(row[h[CFG.HEADERS.DUE]]));

  const lastWatered = formatYmd_(toDate_(row[h[CFG.HEADERS.LAST_WATERED]]));
  const lastFertilized = formatYmd_(toDate_(row[h[CFG.HEADERS.LAST_FERT]]));
  const lastRepotted = formatYmd_(toDate_(row[h[CFG.HEADERS.LAST_REPOT]]));
  const birthday = formatYmd_(toDate_(row[h[CFG.HEADERS.BIRTHDAY]]));

  const notes = String(row[h[CFG.HEADERS.NOTES]] || '');

  const primary = nickname || displayName || name || [genus, species].filter(Boolean).join(' ').trim() || 'Plant';
  const gs = [genus, species].filter(Boolean).join(' ').trim();

  // If Plant ID is purely numeric, treat it as legacy and don't promote it
  const humanPlantId = (plantId && /[A-Za-z]/.test(plantId)) ? plantId : '';

  return {
    ok: true,
    plant: {
      uid: key,
      nickname,
      displayName,
      name,
      primary,
      classification,
      genus,
      species,
      gs,
      plantId,
      humanPlantId,

      location: String(row[h[CFG.HEADERS.LOCATION]] || '').trim(),

      substrate,
      medium,
      potSize,

      birthday,
      notes,

      reminderEnabled,
      everyDays,
      due,

      lastWatered,
      lastFertilized,
      lastRepotted,

      folderId,
      folderUrl: folderId ? `https://drive.google.com/drive/folders/${folderId}` : '',

      qrFileId,
      qrImageUrl: qrFileId ? `https://drive.google.com/uc?export=view&id=${qrFileId}` : '',
      qrUrl,

      careDocId,
      careDocUrl
    }
  };
}

/** Nickname setter used by the UI “Add a nickname?” button */
function plantosSetNickname(uid, nickname) {
  const key = normalizeUid_(uid);
  if (!key) return { ok: false, reason: 'bad_uid' };

  const nn = String(nickname || '').trim();
  if (!nn) return { ok: false, reason: 'empty_nickname' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const inv = found.inv;
  const h = found.h;

  if (CFG.HEADERS.NICKNAME in h) inv.getRange(found.rowIndex, h[CFG.HEADERS.NICKNAME] + 1).setValue(nn);

  // Also set Display Name if it’s blank, so you don’t see numeric junk as primary
  if (CFG.HEADERS.DISPLAY_NAME in h) {
    const dn = String(found.row[h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
    if (!dn) inv.getRange(found.rowIndex, h[CFG.HEADERS.DISPLAY_NAME] + 1).setValue(nn);
  }

  try { ensureHumanFolderNameForUid_(key); } catch (e) {}

  logAction_(key, 'UPDATE', `Set nickname: ${nn}`, null);
  return { ok: true, uid: key, nickname: nn };
}

/** Create plant (basic/advanced). Keeps your sheet columns and adds missing columns without wiping. */
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

    // Keep Plant ID if user supplied, else generate a human one
    const plantId = String(f.plantId || '').trim() || generatePlantId_(uid);

    const meta = {
      uid,
      nickname: f.nickname,
      displayName: f.displayName || f.nickname || f.name || '',
      name: f.name || '',
      plantId,
      classification: f.classification,
      genus: f.genus,
      species: f.species,
      location: f.location,
      substrate: f.substrate,
      medium: f.medium,
      potSize: f.potSize
    };

    const folderData = ensurePlantDriveStructure_(meta);
    const qr = ensurePlantQr_(uid, '', meta.displayName || meta.nickname || meta.name, meta.plantId);
    const care = ensureCareDoc_(uid, folderData.folderId, meta);

    const newRow = new Array(inv.getLastColumn()).fill('');

    // Identity
    if (CFG.HEADERS.UID in h) newRow[h[CFG.HEADERS.UID]] = uid;
    if (CFG.HEADERS.PLANT_ID in h) newRow[h[CFG.HEADERS.PLANT_ID]] = plantId;
    if (CFG.HEADERS.NICKNAME in h) newRow[h[CFG.HEADERS.NICKNAME]] = meta.nickname;
    if (CFG.HEADERS.DISPLAY_NAME in h) newRow[h[CFG.HEADERS.DISPLAY_NAME]] = meta.displayName;
    if (CFG.HEADERS.NAME in h) newRow[h[CFG.HEADERS.NAME]] = meta.name;

    // Taxonomy + placement
    if (CFG.HEADERS.CLASSIFICATION in h) newRow[h[CFG.HEADERS.CLASSIFICATION]] = meta.classification;
    if (CFG.HEADERS.GENUS in h) newRow[h[CFG.HEADERS.GENUS]] = meta.genus;
    if (CFG.HEADERS.SPECIES in h) newRow[h[CFG.HEADERS.SPECIES]] = meta.species;
    if (CFG.HEADERS.LOCATION in h) newRow[h[CFG.HEADERS.LOCATION]] = meta.location;

    // Practical
    if (CFG.HEADERS.SUBSTRATE in h) newRow[h[CFG.HEADERS.SUBSTRATE]] = meta.substrate;
    if (CFG.HEADERS.MEDIUM in h) newRow[h[CFG.HEADERS.MEDIUM]] = meta.medium;
    if (CFG.HEADERS.POT_SIZE in h) newRow[h[CFG.HEADERS.POT_SIZE]] = meta.potSize;
    if (CFG.HEADERS.BIRTHDAY in h) newRow[h[CFG.HEADERS.BIRTHDAY]] = f.birthday || '';

    if (CFG.HEADERS.NOTES in h) newRow[h[CFG.HEADERS.NOTES]] = f.notes || '';

    // Water schedule
    if (CFG.HEADERS.REMINDER_ENABLED in h) newRow[h[CFG.HEADERS.REMINDER_ENABLED]] = !!f.reminderEnabled;
    if (CFG.HEADERS.EVERY in h) newRow[h[CFG.HEADERS.EVERY]] = f.everyDays || '';
    if (f.reminderEnabled && f.everyDays > 0 && (CFG.HEADERS.DUE in h)) {
      const base = startOfDay_(new Date());
      const due = new Date(base); due.setDate(due.getDate() + Number(f.everyDays));
      newRow[h[CFG.HEADERS.DUE]] = due;
    }

    // Drive + QR + care doc
    if (CFG.HEADERS.FOLDER_ID in h) newRow[h[CFG.HEADERS.FOLDER_ID]] = folderData.folderId || '';
    if (CFG.HEADERS.FOLDER_URL in h) newRow[h[CFG.HEADERS.FOLDER_URL]] = folderData.folderId ? `https://drive.google.com/drive/folders/${folderData.folderId}` : '';

    if (CFG.HEADERS.QR_FILE_ID in h) newRow[h[CFG.HEADERS.QR_FILE_ID]] = qr.fileId || '';
    if (CFG.HEADERS.QR_URL in h) newRow[h[CFG.HEADERS.QR_URL]] = qr.url || '';

    if (CFG.HEADERS.CARE_DOC_ID in h) newRow[h[CFG.HEADERS.CARE_DOC_ID]] = care.docId || '';
    if (CFG.HEADERS.CARE_DOC_URL in h) newRow[h[CFG.HEADERS.CARE_DOC_URL]] = care.docUrl || '';

    inv.appendRow(newRow);

    logAction_(uid, 'CREATE', `Created specimen: ${meta.displayName || meta.nickname || meta.plantId || uid}`, meta);
    return { ok: true, uid };
  } finally {
    lock.releaseLock();
  }
}

/** Update plant fields (extended edit) */
function plantosUpdatePlant(uid, patch) {
  const key = normalizeUid_(uid);
  if (!key) return { ok: false, reason: 'bad_uid' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const inv = found.inv;
  const h = found.h;
  const p = patch || {};
  const updates = [];

  function setField_(header, value) {
    if (!(header in h)) return;
    inv.getRange(found.rowIndex, h[header] + 1).setValue(value);
    updates.push(header);
  }

  if (p.nickname != null) setField_(CFG.HEADERS.NICKNAME, String(p.nickname || '').trim());
  if (p.displayName != null) setField_(CFG.HEADERS.DISPLAY_NAME, String(p.displayName || '').trim());
  if (p.name != null) setField_(CFG.HEADERS.NAME, String(p.name || '').trim());
  if (p.plantId != null) setField_(CFG.HEADERS.PLANT_ID, String(p.plantId || '').trim());

  if (p.classification != null) setField_(CFG.HEADERS.CLASSIFICATION, String(p.classification || '').trim());
  if (p.genus != null) setField_(CFG.HEADERS.GENUS, String(p.genus || '').trim());
  if (p.species != null) setField_(CFG.HEADERS.SPECIES, String(p.species || '').trim());

  if (p.location != null) setField_(CFG.HEADERS.LOCATION, String(p.location || '').trim());
  if (p.substrate != null) setField_(CFG.HEADERS.SUBSTRATE, String(p.substrate || '').trim());
  if (p.medium != null) setField_(CFG.HEADERS.MEDIUM, String(p.medium || '').trim());
  if (p.potSize != null) setField_(CFG.HEADERS.POT_SIZE, String(p.potSize || '').trim());

  if (p.birthday != null) setField_(CFG.HEADERS.BIRTHDAY, p.birthday || '');
  if (p.notes != null) setField_(CFG.HEADERS.NOTES, String(p.notes || ''));

  if (p.reminderEnabled != null) setField_(CFG.HEADERS.REMINDER_ENABLED, !!p.reminderEnabled);
  if (p.everyDays != null) setField_(CFG.HEADERS.EVERY, p.everyDays === '' ? '' : Number(p.everyDays));

  if (p.recomputeDue) {
    const now = new Date();
    recomputeNextWaterDue_(key, now);
    updates.push(CFG.HEADERS.DUE);
  }

  if (updates.length) {
    try { ensureHumanFolderNameForUid_(key); } catch (e) {}
    logAction_(key, 'UPDATE', `Updated: ${updates.join(', ')}`, null);
  }
  return { ok: true, updated: updates };
}

/** Quick Log: one call logs multiple actions, updates pot/substrate if included */
function plantosQuickLog(uid, payload) {
  const key = normalizeUid_(uid);
  if (!key) return { ok: false, reason: 'bad_uid' };

  const p = payload || {};
  const water = !!p.water;
  const fertilize = !!p.fertilize;
  const repot = !!p.repot;

  const potSize = String(p.potSize || '').trim();
  const substrate = String(p.substrate || '').trim();
  const medium = String(p.medium || '').trim();
  const notes = String(p.notes || '').trim();

  if (water) logAction_(key, 'WATER', notes ? `QuickLog: ${notes}` : 'QuickLog', null);
  if (fertilize) logAction_(key, 'FERTILIZE', notes ? `QuickLog: ${notes}` : 'QuickLog', null);

  if (repot) {
    const details = [
      potSize ? ('Pot: ' + potSize) : null,
      substrate ? ('Substrate: ' + substrate) : null,
      medium ? ('Medium: ' + medium) : null,
      notes ? ('Notes: ' + notes) : null
    ].filter(Boolean).join(' | ') || 'QuickLog';
    logAction_(key, 'REPOT', details, null);

    // Persist pot/substrate/medium on the plant row too
    const patch = {};
    if (potSize) patch.potSize = potSize;
    if (substrate) patch.substrate = substrate;
    if (medium) patch.medium = medium;
    if (Object.keys(patch).length) plantosUpdatePlant(key, patch);
  }

  if (!water && !fertilize && !repot && notes) {
    logAction_(key, 'NOTE', notes, null);
  }

  // Return fresh plant
  return plantosGetPlantChart(key);
}

/** Batch water from watering dashboard */
function plantosBatchWater(uids) {
  const list = Array.isArray(uids) ? uids : [];
  const trimmed = list.map(normalizeUid_).filter(Boolean).slice(0, CFG.MAX_BATCH_WATER);

  let ok = 0, fail = 0;
  for (const uid of trimmed) {
    try {
      logAction_(uid, 'WATER', 'BatchWater', null);
      ok++;
    } catch (e) {
      fail++;
    }
  }
  return { ok: true, watered: ok, failed: fail };
}

/** Logs */
function plantosGetRecentLog(limit) {
  const lim = Math.max(1, Math.min(200, Number(limit || 25)));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) return [];

  const lastRow = log.getLastRow();
  if (lastRow < 2) return [];

  const start = Math.max(2, lastRow - lim + 1);
  const values = log.getRange(start, 1, lastRow - start + 1, 6).getValues().reverse();

  return values.map(r => ({
    ts: formatYmdTime12h_(toDate_(r[0])),
    plantId: String(r[1] || ''),
    action: String(r[2] || ''),
    details: String(r[3] || ''),
    user: String(r[4] || ''),
    uid: String(r[5] || '')
  }));
}

function plantosGetTimeline(uid, limit) {
  const key = normalizeUid_(uid);
  const lim = Math.max(1, Math.min(200, Number(limit || 30)));
  if (!key) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) return [];

  const lastRow = log.getLastRow();
  if (lastRow < 2) return [];

  const start = Math.max(2, lastRow - 1200 + 1);
  const values = log.getRange(start, 1, lastRow - start + 1, 6).getValues().reverse();

  return values
    .filter(r => String(r[5] || '').trim() === key)
    .slice(0, lim)
    .map(r => ({
      ts: formatYmdTime12h_(toDate_(r[0])),
      plantId: String(r[1] || ''),
      action: String(r[2] || ''),
      details: String(r[3] || ''),
      user: String(r[4] || ''),
      uid: String(r[5] || '')
    }));
}

/** Repair (safe + catches bad folder IDs) */
function plantosRepairUid(uid) {
  ensureSheets_();
  ensureDriveTree_();

  const key = normalizeUid_(uid);
  if (!key) return { ok: false, reason: 'bad_uid' };

  const found = findPlantRowByUid_(key);
  if (!found) return { ok: false, reason: 'not_found' };

  const row = found.row;
  const h = found.h;
  const meta = metaFromRow_(row, h);

  let folderId = String(row[h[CFG.HEADERS.FOLDER_ID]] || '').trim();
  if (!folderId) {
    const fd = ensurePlantDriveStructure_(meta);
    folderId = fd.folderId || '';
    if (folderId) {
      if (CFG.HEADERS.FOLDER_ID in h) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.FOLDER_ID] + 1).setValue(folderId);
      if (CFG.HEADERS.FOLDER_URL in h) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.FOLDER_URL] + 1).setValue(`https://drive.google.com/drive/folders/${folderId}`);
    }
  } else {
    // Rename existing
    try { ensureHumanFolderNameForUid_(key); } catch (e) {}
  }

  const qr = ensurePlantQr_(key, '', meta.displayName || meta.nickname || meta.name, meta.plantId);
  if (qr.fileId) {
    if (CFG.HEADERS.QR_FILE_ID in h) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.QR_FILE_ID] + 1).setValue(qr.fileId);
    if (CFG.HEADERS.QR_URL in h) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.QR_URL] + 1).setValue(qr.url);
  }

  const care = ensureCareDoc_(key, folderId, meta);
  if (care.docId) {
    if (CFG.HEADERS.CARE_DOC_ID in h) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.CARE_DOC_ID] + 1).setValue(care.docId);
    if (CFG.HEADERS.CARE_DOC_URL in h) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.CARE_DOC_URL] + 1).setValue(care.docUrl);
  }

  logAction_(key, 'REPAIR', 'Repaired folder + QR + care doc', meta);
  return { ok: true, uid: key, folderId, qr, care };
}

function plantosRepairAll(limit, startRow) {
  ensureSheets_();
  ensureDriveTree_();

  const invInfo = getInventory_();
  const inv = invInfo.sheet;
  const h = invInfo.h;

  const allData = inv.getDataRange().getValues();
  const start = Math.max(2, Number(startRow || 2));
  const lim = Math.max(1, Math.min(CFG.MAX_BATCH_REPAIR, Number(limit || 200)));

  let checked = 0, fixed = 0;
  for (let r = start; r <= allData.length && checked < lim; r++) {
    checked++;
    const row = allData[r - 1];
    const uid = normalizeUid_(row[h[CFG.HEADERS.UID]]);
    if (!uid) continue;

    try {
      const meta = metaFromRow_(row, h);

      // Folder
      let folderId = String(row[h[CFG.HEADERS.FOLDER_ID]] || '').trim();
      if (!folderId) {
        const fd = ensurePlantDriveStructure_(meta);
        folderId = fd.folderId || '';
        if (folderId) {
          inv.getRange(r, h[CFG.HEADERS.FOLDER_ID] + 1).setValue(folderId);
          if (CFG.HEADERS.FOLDER_URL in h) inv.getRange(r, h[CFG.HEADERS.FOLDER_URL] + 1).setValue(`https://drive.google.com/drive/folders/${folderId}`);
          fixed++;
        }
      } else {
        try { ensureHumanFolderNameForUid_(uid); } catch (e) {}
      }

      // QR
      const qrId = String(row[h[CFG.HEADERS.QR_FILE_ID]] || '').trim();
      if (!qrId) {
        const qr = ensurePlantQr_(uid, '', meta.displayName || meta.nickname || meta.name, meta.plantId);
        if (qr.fileId) {
          inv.getRange(r, h[CFG.HEADERS.QR_FILE_ID] + 1).setValue(qr.fileId);
          inv.getRange(r, h[CFG.HEADERS.QR_URL] + 1).setValue(qr.url);
          fixed++;
        }
      }

      // Care doc
      const careId = (CFG.HEADERS.CARE_DOC_ID in h) ? String(row[h[CFG.HEADERS.CARE_DOC_ID]] || '').trim() : '';
      if (!careId) {
        const folderId2 = String(inv.getRange(r, h[CFG.HEADERS.FOLDER_ID] + 1).getValue() || '').trim();
        if (folderId2) {
          const care = ensureCareDoc_(uid, folderId2, meta);
          if (care.docId) {
            inv.getRange(r, h[CFG.HEADERS.CARE_DOC_ID] + 1).setValue(care.docId);
            inv.getRange(r, h[CFG.HEADERS.CARE_DOC_URL] + 1).setValue(care.docUrl);
            fixed++;
          }
        }
      }
    } catch (e) {
      // swallow and continue
    }
  }

  const nextRow = start + checked;
  logAction_('SYSTEM', 'REPAIR_ALL', `Checked=${checked}, Fixed=${fixed}, NextRow=${nextRow}`, null);
  return { ok: true, checked, fixed, nextRow };
}

/** QR helpers */
function plantosRebuildPlantQrs(limit, overrideUrl) {
  ensureSheets_();
  ensureDriveTree_();

  const invInfo = getInventory_();
  const inv = invInfo.sheet;
  const h = invInfo.h;
  const allData = inv.getDataRange().getValues();

  const lim = Math.max(1, Math.min(5000, Number(limit || 300)));
  const baseUrl = String(overrideUrl || '').trim() || getBaseUrl_();

  let count = 0;
  for (let r = 2; r <= allData.length && count < lim; r++) {
    const row = allData[r - 1];
    const uid = normalizeUid_(row[h[CFG.HEADERS.UID]]);
    if (!uid) continue;

    const meta = metaFromRow_(row, h);
    const qr = ensurePlantQr_(uid, baseUrl, meta.displayName || meta.nickname || meta.name, meta.plantId);
    if (qr.fileId) {
      inv.getRange(r, h[CFG.HEADERS.QR_FILE_ID] + 1).setValue(qr.fileId);
      inv.getRange(r, h[CFG.HEADERS.QR_URL] + 1).setValue(qr.url);
    }
    count++;
  }

  logAction_('SYSTEM', 'QR_PLANT_REBUILD', `Plant QRs rebuilt: ${count}`, null);
  return { ok: true, count };
}

/** ===================== LOGGING ENGINE ===================== **/

function logAction_(uid, action, details, metaMaybe) {
  ensureSheets_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) throw new Error('Missing log sheet: ' + CFG.LOG_SHEET);

  let userEmail = '';
  try { userEmail = Session.getActiveUser().getEmail(); } catch (e) {}

  const key = String(uid || '').trim();
  const normalizedUid = normalizeUid_(key);

  const plantId = normalizedUid ? getPlantIdByUid_(normalizedUid) : '';
  log.appendRow([new Date(), plantId, String(action || ''), String(details || ''), userEmail, normalizedUid || key]);

  if (normalizedUid) {
    const found = findPlantRowByUid_(normalizedUid);
    const row = found ? found.row : null;
    const h = found ? found.h : null;
    const meta = metaMaybe || (row && h ? metaFromRow_(row, h) : { uid: normalizedUid, displayName: normalizedUid, plantId });

    if (action === 'WATER') {
      const now = new Date();
      updateInventoryDate_(normalizedUid, CFG.HEADERS.LAST_WATERED, now);
      recomputeNextWaterDue_(normalizedUid, now);
    }
    if (action === 'FERTILIZE') updateInventoryDate_(normalizedUid, CFG.HEADERS.LAST_FERT, new Date());
    if (action === 'REPOT') updateInventoryDate_(normalizedUid, CFG.HEADERS.LAST_REPOT, new Date());
    if (action === 'NOTE' && details) updateInventoryText_(normalizedUid, CFG.HEADERS.NOTES, details);

    try { appendToCareDoc_(normalizedUid, action, details, meta); } catch (e) {}
  }
}

/** ===================== SHEETS + SETTINGS HELPERS ===================== **/

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let inv = ss.getSheetByName(CFG.INVENTORY_SHEET);
  if (!inv) inv = ss.insertSheet(CFG.INVENTORY_SHEET);

  let log = ss.getSheetByName(CFG.LOG_SHEET);
  if (!log) log = ss.insertSheet(CFG.LOG_SHEET);

  let settings = ss.getSheetByName(CFG.SETTINGS_SHEET);
  if (!settings) settings = ss.insertSheet(CFG.SETTINGS_SHEET);

  // Ensure Plant Log header (6 cols)
  if (log.getLastRow() === 0) log.appendRow(['Timestamp', 'Plant ID', 'Action', 'Details', 'User', 'Plant UID']);
  if (log.getLastRow() >= 1) {
    const h0 = String(log.getRange(1, 1, 1, 1).getValues()[0][0] || '');
    if (h0 !== 'Timestamp') {
      log.insertRowBefore(1);
      log.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Plant ID', 'Action', 'Details', 'User', 'Plant UID']]);
    }
  }

  if (settings.getLastRow() === 0) settings.appendRow(['KEY', 'VALUE']);

  // Ensure inventory missing columns are appended, never wiped
  ensureInventoryColumns_(inv);
}

function ensureInventoryColumns_(inv) {
  const lastCol = Math.max(1, inv.getLastColumn());
  const headers = inv.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const h = getHeaderMap_(headers);

  const required = Object.values(CFG.HEADERS);
  const missing = required.filter(col => !(col in h));

  if (headers.filter(Boolean).length === 0) {
    inv.clear();
    inv.getRange(1, 1, 1, required.length).setValues([required]);
    return;
  }

  if (missing.length) {
    inv.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  }
}

function getInventory_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CFG.INVENTORY_SHEET);
  if (!sheet) throw new Error('Missing inventory sheet: ' + CFG.INVENTORY_SHEET);

  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const h = getHeaderMap_(headers);

  const must = [CFG.HEADERS.UID, CFG.HEADERS.LOCATION, CFG.HEADERS.FOLDER_ID, CFG.HEADERS.QR_FILE_ID, CFG.HEADERS.QR_URL];
  for (const m of must) {
    if (!(m in h)) throw new Error(`Missing required column "${m}" in "${CFG.INVENTORY_SHEET}".`);
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
  const saved = getSetting_('ACTIVE_WEBAPP_URL');
  if (saved) return saved;
  try {
    const u = ScriptApp.getService().getUrl();
    if (u) return u;
  } catch (e) {}
  return '';
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

function normalizeUid_(v) {
  if (v == null) return '';
  const s = String(v).trim();
  if (/^\d{8}$/.test(s)) return s;

  // If stored as number, String() already worked. If it’s like 8.23433E+07 we cannot safely recover.
  // But your sheet shows plain 8-digit values, so this is fine.
  return '';
}

function resolveAnyToUid_(raw) {
  const input = String(raw || '').trim();
  if (!input) return '';
  if (/^\d{8}$/.test(input)) return input;

  const invInfo = getInventory_();
  const data = invInfo.sheet.getDataRange().getValues();
  const h = invInfo.h;
  const keyLower = input.toLowerCase();

  // Plant ID exact
  if (CFG.HEADERS.PLANT_ID in h) {
    for (let i = 1; i < data.length; i++) {
      const pid = String(data[i][h[CFG.HEADERS.PLANT_ID]] || '').trim();
      if (pid && pid.toLowerCase() === keyLower) return normalizeUid_(data[i][h[CFG.HEADERS.UID]]);
    }
  }

  // Nickname exact
  if (CFG.HEADERS.NICKNAME in h) {
    for (let i = 1; i < data.length; i++) {
      const nn = String(data[i][h[CFG.HEADERS.NICKNAME]] || '').trim();
      if (nn && nn.toLowerCase() === keyLower) return normalizeUid_(data[i][h[CFG.HEADERS.UID]]);
    }
  }

  // Display Name exact
  if (CFG.HEADERS.DISPLAY_NAME in h) {
    for (let i = 1; i < data.length; i++) {
      const dn = String(data[i][h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
      if (dn && dn.toLowerCase() === keyLower) return normalizeUid_(data[i][h[CFG.HEADERS.UID]]);
    }
  }

  // Partial (nickname/display/name)
  for (let i = 1; i < data.length; i++) {
    const dn = (CFG.HEADERS.DISPLAY_NAME in h) ? String(data[i][h[CFG.HEADERS.DISPLAY_NAME]] || '').trim() : '';
    const nn = (CFG.HEADERS.NICKNAME in h) ? String(data[i][h[CFG.HEADERS.NICKNAME]] || '').trim() : '';
    const nm = (CFG.HEADERS.NAME in h) ? String(data[i][h[CFG.HEADERS.NAME]] || '').trim() : '';
    const blob = (dn + ' ' + nn + ' ' + nm).toLowerCase();
    if (blob.includes(keyLower)) return normalizeUid_(data[i][h[CFG.HEADERS.UID]]);
  }

  return '';
}

function findPlantRowByUid_(uid) {
  const key = normalizeUid_(uid);
  if (!key) return null;

  const invInfo = getInventory_();
  const inv = invInfo.sheet;
  const data = inv.getDataRange().getValues();
  const h = invInfo.h;

  for (let i = 1; i < data.length; i++) {
    if (normalizeUid_(data[i][h[CFG.HEADERS.UID]]) === key) {
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
      const s = normalizeUid_(v[0]);
      if (s) existing.add(s);
    });
  }

  let uid = '';
  do {
    uid = Math.floor(10000000 + Math.random() * 90000000).toString();
  } while (existing.has(uid));

  return uid;
}

function generatePlantId_(uid) {
  const s = String(uid || '').trim();
  if (!/^\d{8}$/.test(s)) return 'PL-' + s.slice(-6);
  const base36 = Number(s).toString(36).toUpperCase();
  return 'PL-' + base36.slice(-6);
}

function getPlantIdByUid_(uid) {
  const found = findPlantRowByUid_(uid);
  if (!found) return '';
  if (!(CFG.HEADERS.PLANT_ID in found.h)) return '';
  return String(found.row[found.h[CFG.HEADERS.PLANT_ID]] || '').trim();
}

/** ===================== LABELING HELPERS ===================== **/

function taxonKey_(row, h) {
  const classification = String(row[h[CFG.HEADERS.CLASSIFICATION]] || '').trim().toLowerCase();
  const genus = String(row[h[CFG.HEADERS.GENUS]] || '').trim().toLowerCase();
  const species = String(row[h[CFG.HEADERS.SPECIES]] || '').trim().toLowerCase();
  return [classification, genus, species].join('|');
}

function taxonLabel_(classification, genus, species) {
  const gs = [genus, species].filter(Boolean).join(' ').trim();
  if (gs) return gs;
  return String(classification || 'Plant').trim() || 'Plant';
}

function computeDisplayLabel_(row, h, fallbackTax) {
  const nickname = String(row[h[CFG.HEADERS.NICKNAME]] || '').trim();
  const displayName = String(row[h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
  const name = String(row[h[CFG.HEADERS.NAME]] || '').trim();

  const classification = String(row[h[CFG.HEADERS.CLASSIFICATION]] || '').trim();
  const genus = String(row[h[CFG.HEADERS.GENUS]] || '').trim();
  const species = String(row[h[CFG.HEADERS.SPECIES]] || '').trim();

  const primary = nickname || displayName || name || (fallbackTax ? taxonLabel_(classification, genus, species) : 'Plant');
  const gs = [genus, species].filter(Boolean).join(' ').trim();
  const secondary = nickname ? (gs || classification || '') : (classification || gs || '');
  return { primary, secondary };
}

function metaFromRow_(row, h) {
  const uid = normalizeUid_(row[h[CFG.HEADERS.UID]]);
  const nickname = String(row[h[CFG.HEADERS.NICKNAME]] || '').trim();
  const displayName = String(row[h[CFG.HEADERS.DISPLAY_NAME]] || '').trim();
  const name = String(row[h[CFG.HEADERS.NAME]] || '').trim();
  const plantId = String(row[h[CFG.HEADERS.PLANT_ID]] || '').trim();

  return {
    uid,
    nickname,
    displayName: displayName || nickname || name || '',
    name,
    plantId,
    classification: String(row[h[CFG.HEADERS.CLASSIFICATION]] || '').trim(),
    genus: String(row[h[CFG.HEADERS.GENUS]] || '').trim(),
    species: String(row[h[CFG.HEADERS.SPECIES]] || '').trim(),
    location: String(row[h[CFG.HEADERS.LOCATION]] || '').trim(),
    substrate: String(row[h[CFG.HEADERS.SUBSTRATE]] || '').trim(),
    medium: String(row[h[CFG.HEADERS.MEDIUM]] || '').trim(),
    potSize: String(row[h[CFG.HEADERS.POT_SIZE]] || '').trim()
  };
}

function normalizeForm_(form) {
  const f = form || {};
  return {
    displayName: String(f.displayName || '').trim(),
    nickname: String(f.nickname || '').trim(),
    name: String(f.name || '').trim(),
    plantId: String(f.plantId || '').trim(),

    classification: String(f.classification || '').trim(),
    genus: String(f.genus || '').trim(),
    species: String(f.species || '').trim(),

    location: String(f.location || '').trim(),

    substrate: String(f.substrate || '').trim(),
    medium: String(f.medium || '').trim(),
    potSize: String(f.potSize || '').trim(),

    birthday: String(f.birthday || '').trim(),
    notes: String(f.notes || ''),

    reminderEnabled: (String(f.reminderEnabled || '').toLowerCase() === 'true') || f.reminderEnabled === true,
    everyDays: (f.everyDays === '' || f.everyDays == null) ? 0 : Math.max(0, Math.floor(Number(f.everyDays)))
  };
}

/** ===================== DRIVE HELPERS ===================== **/

function ensureDriveTree_() {
  const cachedRootId = getSetting_('DRIVE_ROOT_ID');

  const root = findOrCreateFolder_(DriveApp.getRootFolder(), CFG.ROOT_NAME);
  const plants = findOrCreateFolder_(root, CFG.DRIVE.ROOT_FOLDERS.PLANTS);
  const locations = findOrCreateFolder_(plants, CFG.DRIVE.PLANTS_FOLDERS.LOCATIONS);

  const qr = findOrCreateFolder_(root, CFG.DRIVE.ROOT_FOLDERS.QR);
  const qrSystem = findOrCreateFolder_(qr, CFG.DRIVE.QR_FOLDERS.SYSTEM);
  const qrLocations = findOrCreateFolder_(qr, CFG.DRIVE.QR_FOLDERS.LOCATIONS);
  const qrPlants = findOrCreateFolder_(qr, CFG.DRIVE.QR_FOLDERS.PLANTS);

  if (root.getId() !== cachedRootId) {
    setSetting_('DRIVE_ROOT_ID', root.getId());
    setSetting_('DRIVE_PLANTS_ID', plants.getId());
    setSetting_('DRIVE_LOCATIONS_ID', locations.getId());
    setSetting_('DRIVE_QR_ID', qr.getId());
    setSetting_('DRIVE_QR_SYSTEM_ID', qrSystem.getId());
    setSetting_('DRIVE_QR_LOCATIONS_ID', qrLocations.getId());
    setSetting_('DRIVE_QR_PLANTS_ID', qrPlants.getId());
  }

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

function plantFolderName_(meta) {
  const primary = String(meta.nickname || meta.displayName || meta.name || 'Plant').trim() || 'Plant';
  const cls = String(meta.classification || '').trim();
  const genus = String(meta.genus || '').trim();
  const species = String(meta.species || '').trim();

  const gs = [genus, species].filter(Boolean).join(' ').trim();
  const tax = [cls, gs].filter(Boolean).join(' — ');

  const pid = String(meta.plantId || '').trim();
  const parts = [];
  parts.push(primary);
  if (tax) parts.push(tax);

  // Only include Plant ID if it’s actually “human” (has letters)
  if (pid && /[A-Za-z]/.test(pid)) parts.push(pid);

  // UID always last
  parts.push(meta.uid);
  return safeFolderName_(parts.join(' — '));
}

function ensurePlantDriveStructure_(meta) {
  const tree = ensureDriveTree_();
  const locationsFolder = DriveApp.getFolderById(tree.locationsId);

  const locName = String(meta.location || '').trim() || 'Unsorted';
  const locFolder = findOrCreateFolder_(locationsFolder, locName);

  let plantFolder = findChildFolderByUid_(locFolder, meta.uid);
  if (!plantFolder) {
    plantFolder = locFolder.createFolder(plantFolderName_(meta));
  } else {
    try { plantFolder.setName(plantFolderName_(meta)); } catch (e) {}
  }

  CFG.DRIVE.PLANT_SUBFOLDERS.forEach(sub => findOrCreateFolder_(plantFolder, sub));
  return { folderId: plantFolder.getId() };
}

function ensureHumanFolderNameForUid_(uid) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;

  const folderId = String(found.row[found.h[CFG.HEADERS.FOLDER_ID]] || '').trim();
  if (!folderId) return;

  const meta = metaFromRow_(found.row, found.h);
  try {
    DriveApp.getFolderById(folderId).setName(plantFolderName_(meta));
  } catch (e) {}
}

function qrFileName_(displayName, plantId, uid) {
  const a = safeFileName_(String(displayName || '').trim());
  const b = safeFileName_(String(plantId || '').trim());
  if (a && b) return `${a}_${b}_${uid}.png`;
  if (a) return `${a}_${uid}.png`;
  if (b) return `${b}_${uid}.png`;
  return `${uid}.png`;
}

function ensurePlantQr_(uid, overrideUrl, displayName, plantId) {
  const tree = ensureDriveTree_();
  const qrPlants = DriveApp.getFolderById(tree.qrPlantsId);

  const baseUrl = String(overrideUrl || '').trim() || getBaseUrl_();
  const url = baseUrl ? `${baseUrl}?mode=plant&uid=${encodeURIComponent(uid)}` : String(uid);

  const fileName = qrFileName_(displayName, plantId, uid);
  const file = upsertQrPng_(qrPlants, fileName, url);
  return { url, fileId: file ? file.getId() : '', fileName };
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

function ensureCareDoc_(uid, folderId, meta) {
  if (!folderId) return { docId: '', docUrl: '' };

  // If sheet already has care doc id and it exists, reuse
  const found = findPlantRowByUid_(uid);
  if (found && (CFG.HEADERS.CARE_DOC_ID in found.h)) {
    const existingId = String(found.row[found.h[CFG.HEADERS.CARE_DOC_ID]] || '').trim();
    if (existingId) {
      try {
        DriveApp.getFileById(existingId);
        return { docId: existingId, docUrl: `https://docs.google.com/document/d/${existingId}/edit` };
      } catch (e) {}
    }
  }

  const plantFolder = DriveApp.getFolderById(folderId);
  const careFolder = findOrCreateFolder_(plantFolder, 'Care Notes');

  const title = `Care Log — ${meta.nickname || meta.displayName || meta.name || meta.plantId || 'Plant'} — ${meta.genus || ''} ${meta.species || ''} — ${uid}`.replace(/\s+/g, ' ').trim();
  const doc = DocumentApp.create(title);
  const docId = doc.getId();

  const file = DriveApp.getFileById(docId);
  careFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const body = doc.getBody();
  body.appendParagraph('PlantOS Care Log').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`Nickname: ${meta.nickname || ''}`);
  body.appendParagraph(`Display: ${meta.displayName || ''}`);
  body.appendParagraph(`Plant ID: ${meta.plantId || ''}`);
  body.appendParagraph(`UID: ${uid}`);
  body.appendParagraph(`Taxonomy: ${(meta.classification || '')} / ${(meta.genus || '')} ${(meta.species || '')}`.trim());
  body.appendParagraph('---');
  doc.saveAndClose();

  return { docId, docUrl: `https://docs.google.com/document/d/${docId}/edit` };
}

function appendToCareDoc_(uid, action, details, meta) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;

  const h = found.h;
  const row = found.row;
  const folderId = String(row[h[CFG.HEADERS.FOLDER_ID]] || '').trim();
  if (!folderId) return;

  const care = ensureCareDoc_(uid, folderId, meta);

  // Persist care doc back to sheet if missing
  if (care.docId) {
    if (CFG.HEADERS.CARE_DOC_ID in h) {
      const existing = String(row[h[CFG.HEADERS.CARE_DOC_ID]] || '').trim();
      if (!existing) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.CARE_DOC_ID] + 1).setValue(care.docId);
    }
    if (CFG.HEADERS.CARE_DOC_URL in h) {
      const existing = String(row[h[CFG.HEADERS.CARE_DOC_URL]] || '').trim();
      if (!existing) found.inv.getRange(found.rowIndex, h[CFG.HEADERS.CARE_DOC_URL] + 1).setValue(care.docUrl);
    }
  }

  const doc = DocumentApp.openById(care.docId);
  const body = doc.getBody();
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a'); // 12-hour
  const line = `[${stamp}] ${action}${details ? ' — ' + details : ''}`;
  body.appendParagraph(line);
  doc.saveAndClose();
}

function findOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function findChildFolderByUid_(parent, uid) {
  const it = parent.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (String(f.getName() || '').includes(uid)) return f;
  }
  return null;
}

/** ===================== INVENTORY MUTATORS ===================== **/

function updateInventoryDate_(uid, header, date) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;
  if (!(header in found.h)) return;
  found.inv.getRange(found.rowIndex, found.h[header] + 1).setValue(date);
}

function updateInventoryText_(uid, header, text) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;
  if (!(header in found.h)) return;
  found.inv.getRange(found.rowIndex, found.h[header] + 1).setValue(text);
}

function recomputeNextWaterDue_(uid, baseDate) {
  const found = findPlantRowByUid_(uid);
  if (!found) return;

  const every = Number(found.row[found.h[CFG.HEADERS.EVERY]] || 0);
  const enabled = found.row[found.h[CFG.HEADERS.REMINDER_ENABLED]] === true;
  if (!enabled || !every || every < 1) return;

  const d = startOfDay_(new Date(baseDate || new Date()));
  d.setDate(d.getDate() + every);
  updateInventoryDate_(uid, CFG.HEADERS.DUE, d);
}

/** ===================== DATE HELPERS ===================== **/

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

function formatYmdTime12h_(d) {
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd hh:mm a');
}

function safeFileName_(name) {
  const n = String(name || '').trim() || ('file_' + Date.now());
  return n.replace(/[\/\\:*?"<>|]+/g, '_').replace(/\s+/g, ' ').trim();
}

function safeFolderName_(name) {
  const n = String(name || '').trim() || ('Plant_' + Date.now());
  return n.replace(/[\/\\:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 180);
}

/** ===================== OPTIONAL SHEETS MENU (tiny but useful) ===================== **/
function onOpen(e) {
  try { plantosBuildSheetsMenu_(); } catch (err) {}
}

function plantosBuildSheetsMenu_() {
  const ui = SpreadsheetApp.getUi();
  const m = ui.createMenu('PlantOS');
  m.addItem('Initialize (Drive tree)', 'plantosMenu_Init_');
  m.addItem('Repair All (chunk)', 'plantosMenu_RepairAllChunk_');
  m.addToUi();
}

function plantosMenu_Init_() {
  const ui = SpreadsheetApp.getUi();
  const cur = getBaseUrl_() || '';
  const p = ui.prompt('PlantOS URL', 'Paste your deployed Web App URL (used for QRs). Leave blank to keep current.\n\nCurrent:\n' + cur, ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const entered = String(p.getResponseText() || '').trim();
  if (entered) setSetting_('ACTIVE_WEBAPP_URL', entered);
  const res = plantosInit();
  ui.alert(res && res.ok ? 'Initialized ✅' : 'Init failed');
}

function plantosMenu_RepairAllChunk_() {
  const ui = SpreadsheetApp.getUi();
  const start = Number(getSetting_('REPAIR_CURSOR') || '2');
  const res = plantosRepairAll(200, start);
  if (res && res.ok) {
    setSetting_('REPAIR_CURSOR', String(res.nextRow || (start + 200)));
    ui.alert(`Repair chunk ✅\nChecked=${res.checked}\nFixed=${res.fixed}\nNextRow=${res.nextRow}`);
  } else {
    ui.alert('Repair failed');
  }
}
