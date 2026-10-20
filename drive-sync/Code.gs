/**
 * Keepsake: Google Drive sync service (Google Apps Script)
 *
 * Folders it creates in your Drive:
 *   Keepsake/
 *     Photos/2026/10 October/...         your original photos, by month
 *     Data/keepsake.json                 dates, reminders, journal, vault, habits
 *     Data/Thumbnails/...                small previews used by the app
 *     Backups/keepsake-2026-10-20.json   one copy per day, last 30 kept
 *
 * Setup:
 *   1. Set SECRET_PASSKEY below to the password you use in Keepsake.
 *   2. Run the `setup` function once from the editor and allow Drive access.
 *   3. Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy.
 *      Editing the existing deployment keeps the same web app URL.
 */
const SECRET_PASSKEY = 'CHANGE_ME';

const ROOT_NAME = 'Keepsake';
const DATA_FILE = 'keepsake.json';
const KEEP_BACKUPS = 30;
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const WRITE_ACTIONS = ['loadData', 'saveData', 'uploadPhoto', 'deletePhoto'];

function setup() {
  rootFolder();
  dataFolder();
  thumbsFolder();
  photosFolder();
  backupsFolder();
  Logger.log('Keepsake folders are ready: ' + rootFolder().getUrl());
}

function doGet() {
  return out({ ok: true, app: 'Keepsake' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return out({ error: 'Bad request' });
  }
  if (!req || req.secret !== SECRET_PASSKEY) return out({ error: 'Invalid passkey' });

  const lock = WRITE_ACTIONS.indexOf(req.action) >= 0 ? LockService.getScriptLock() : null;
  if (lock) lock.waitLock(30000);
  try {
    switch (req.action) {
      case 'status': return out(status());
      case 'loadData': return out(loadData());
      case 'saveData': return out(saveData(req.data, req.baseRev));
      case 'uploadPhoto': return out(uploadPhoto(req));
      case 'deletePhoto': return out(deletePhoto(req.fileId));
      case 'getPhoto': return out(getPhoto(req.fileId, req.size));
      default: return out({ error: 'Unknown action: ' + req.action });
    }
  } catch (err) {
    return out({ error: String((err && err.message) || err) });
  } finally {
    if (lock) lock.releaseLock();
  }
}

// ---------- Actions ----------

// Cheap check the app polls: the data revision plus a fingerprint of the photo files,
// so photos added to or removed from Drive directly are noticed too.
function status() {
  return { rev: getRev(), photos: photoHash(livePhotos()) };
}

function loadData() {
  const data = readData();
  const files = livePhotos();
  let rev = getRev();
  if (data && reconcile(data, files)) {
    writeData(data);
    rev = bumpRev();
  }
  return { data: data, rev: rev, photos: photoHash(files) };
}

function saveData(data, baseRev) {
  if (!data || typeof data !== 'object') return { error: 'No data' };
  const rev = getRev();
  if (typeof baseRev === 'number' && baseRev < rev) {
    const current = readData();
    if (current) return { status: 'conflict', rev: rev, data: current };
  }
  (data.photos || []).forEach(function (p) { delete p.url; delete p.driveUrl; });
  dailyBackup();
  writeData(data);
  return { status: 'saved', rev: bumpRev() };
}

function uploadPhoto(req) {
  const m = /^data:([^;]+);base64,(.*)$/.exec(req.dataUrl || '');
  if (!m) return { error: 'Invalid image' };
  const photo = req.photo || {};
  const tz = Session.getScriptTimeZone();
  const when = new Date(photo.addedAt || Date.now());
  const year = child(photosFolder(), Utilities.formatDate(when, tz, 'yyyy'));
  const monthNum = Utilities.formatDate(when, tz, 'MM');
  const month = child(year, monthNum + ' ' + MONTHS[Number(monthNum) - 1]);
  const name = String(req.fileName || 'photo.jpg').replace(/[\\/:*?"<>|]/g, '_');
  const file = month.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name));

  const t = /^data:([^;]+);base64,(.*)$/.exec(req.thumbDataUrl || '');
  if (t) thumbsFolder().createFile(Utilities.newBlob(Utilities.base64Decode(t[2]), t[1], file.getId() + '.jpg'));

  const data = readData() || {};
  data.photos = Array.isArray(data.photos) ? data.photos : [];
  data.photos = data.photos.filter(function (p) { return String(p.id) !== String(photo.id); });
  data.photos.unshift({
    id: photo.id,
    fileId: file.getId(),
    caption: photo.caption || name.replace(/\.[^.]+$/, ''),
    date: photo.date || Utilities.formatDate(when, tz, 'M/d/yyyy'),
    addedAt: when.getTime()
  });
  writeData(data);
  return { fileId: file.getId(), rev: bumpRev() };
}

function deletePhoto(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* already gone */ }
  trashThumb(fileId);
  const data = readData();
  if (data && Array.isArray(data.photos)) {
    data.deleted = data.deleted || {};
    data.photos = data.photos.filter(function (p) {
      if (p.fileId !== fileId) return true;
      data.deleted[String(p.id)] = Date.now();
      return false;
    });
    writeData(data);
  }
  return { status: 'deleted', rev: bumpRev() };
}

function getPhoto(fileId, size) {
  if (size === 'thumb') {
    const t = findFile(thumbsFolder(), fileId + '.jpg');
    if (t) return { dataUrl: toDataUrl(t.getBlob()) };
    const th = DriveApp.getFileById(fileId).getThumbnail();
    if (th) return { dataUrl: toDataUrl(th) };
  }
  return { dataUrl: toDataUrl(DriveApp.getFileById(fileId).getBlob()) };
}

// ---------- Drive helpers ----------

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function props() {
  return PropertiesService.getScriptProperties();
}
function getRev() {
  return Number(props().getProperty('rev') || 0);
}
function bumpRev() {
  const rev = getRev() + 1;
  props().setProperty('rev', String(rev));
  return rev;
}

function child(parent, name) {
  const it = parent.getFoldersByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return parent.createFolder(name);
}
function rootFolder() {
  const id = props().getProperty('rootId');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* recreate below */ }
  }
  const f = child(DriveApp.getRootFolder(), ROOT_NAME);
  props().setProperty('rootId', f.getId());
  return f;
}
function dataFolder() { return child(rootFolder(), 'Data'); }
function thumbsFolder() { return child(dataFolder(), 'Thumbnails'); }
function photosFolder() { return child(rootFolder(), 'Photos'); }
function backupsFolder() { return child(rootFolder(), 'Backups'); }

function findFile(folder, name) {
  const it = folder.getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;
  }
  return null;
}

function readData() {
  const f = findFile(dataFolder(), DATA_FILE);
  return f ? JSON.parse(f.getBlob().getDataAsString()) : null;
}
function writeData(data) {
  const json = JSON.stringify(data);
  const folder = dataFolder();
  const f = findFile(folder, DATA_FILE);
  if (f) f.setContent(json);
  else folder.createFile(DATA_FILE, json, 'application/json');
}

// Keeps a copy of the data as it was at the start of each day.
function dailyBackup() {
  const f = findFile(dataFolder(), DATA_FILE);
  if (!f) return;
  const name = 'keepsake-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.json';
  const folder = backupsFolder();
  if (findFile(folder, name)) return;
  f.makeCopy(name, folder);
  const all = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const x = it.next();
    if (!x.isTrashed()) all.push(x);
  }
  all.sort(function (a, b) { return b.getName().localeCompare(a.getName()); });
  all.slice(KEEP_BACKUPS).forEach(function (x) { x.setTrashed(true); });
}

function livePhotos() {
  const files = {};
  (function walk(folder) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (!f.isTrashed() && /^image\//.test(f.getMimeType())) files[f.getId()] = f;
    }
    const sub = folder.getFolders();
    while (sub.hasNext()) {
      const s = sub.next();
      if (!s.isTrashed()) walk(s);
    }
  })(photosFolder());
  return files;
}

function photoHash(files) {
  const ids = Object.keys(files).sort().join(',');
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, ids));
}

// Makes the data match the Photos folder: photos deleted in Drive disappear from the app,
// and images you drop into the Photos folder show up in the app.
function reconcile(data, files) {
  data.photos = Array.isArray(data.photos) ? data.photos : [];
  data.deleted = data.deleted || {};
  const tz = Session.getScriptTimeZone();
  const known = {};
  let changed = false;

  data.photos = data.photos.filter(function (p) {
    if (!p.fileId) return true;
    if (files[p.fileId]) {
      known[p.fileId] = true;
      return true;
    }
    data.deleted[String(p.id)] = Date.now();
    trashThumb(p.fileId);
    changed = true;
    return false;
  });

  Object.keys(files).forEach(function (id) {
    if (known[id]) return;
    const f = files[id];
    const created = f.getDateCreated();
    data.photos.push({
      id: 'd_' + id,
      fileId: id,
      caption: f.getName().replace(/\.[^.]+$/, ''),
      date: Utilities.formatDate(created, tz, 'M/d/yyyy'),
      addedAt: created.getTime()
    });
    changed = true;
  });
  return changed;
}

function trashThumb(fileId) {
  const f = findFile(thumbsFolder(), fileId + '.jpg');
  if (f) f.setTrashed(true);
}

function toDataUrl(blob) {
  return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
}
