/**
 * Hermetic AI — account setup: Apps Script backend
 *
 * Deploy: Extensions > Apps Script, paste this in. Paste appsscript.json too
 * (Project Settings > "Show appsscript.json manifest file in editor", then open
 * it from the file list). Run authorize() once from the editor and accept the
 * Drive prompt, then Deploy > New deployment > Web app, "Execute as: Me", "Who
 * has access: Anyone". Copy the /exec URL into the app's Settings > Google
 * Sheet sync. On later edits use Deploy > Manage deployments > New version,
 * which keeps the same /exec URL.
 *
 * SCOPES
 * A deployment runs with the scopes granted when it was last authorized, not
 * the ones its current code needs. Adding the Drive upload to a script that
 * had only ever read a Sheet therefore fails until you re-consent -- which is
 * what authorize() is for.
 *
 * The manifest is half of that: when appsscript.json lists oauthScopes, that
 * list is the whole request, overriding what Apps Script would infer from the
 * code. So a project still carrying the old Sheets-only manifest re-consents to
 * Sheets only, and DriveApp keeps failing however many times you redeploy.
 * Updating the .gs without the manifest is the usual reason this error sticks.
 * Hit ?action=diag on the deployed URL to see which half is missing.
 *
 * WHY getAll EXISTS
 * Each /exec call costs roughly a second, and same-user calls run one at a
 * time. Reading key-by-key meant a cold page load spent 3 + one-per-venue
 * serial round trips before rendering anything. getAll returns every
 * hermetic: key in a single bulk sheet read, so a cold load costs one call.
 *
 * CONCURRENT EDITORS
 * Several people use this at once, each on a live customer call. Two things
 * here make that safe. setKey_ takes a script lock so writes cannot interleave,
 * and -- separately, because the lock does nothing about it -- it honours an
 * ifRev guard, refusing a write whose base revision has moved on and handing
 * the current value back so the client can merge. getAll takes a `since` so
 * everyone can poll for other people's changes without dragging the whole
 * sheet down every time.
 *
 * SHEET LAYOUT
 * A tab named "kv": column A is the key, column B onward is the JSON value.
 * Values longer than one cell can hold are split across B, C, D... and
 * rejoined on read, so a long source_material paste is never truncated.
 */

var SHEET_NAME = 'kv';
var CELL_LIMIT = 45000;   // Sheets caps a cell at 50k characters; leave headroom.
var DEFAULT_ROOT = 'ONBOARDING - Custom Prompt Uploads';

var REAUTH_HINT =
  'This deployment cannot write to Drive. Read-only Drive access is not enough ' +
  '-- uploads create folders and files, which needs the full ' +
  'https://www.googleapis.com/auth/drive scope. (1) Check appsscript.json lists ' +
  'that exact scope, not drive.readonly or drive.file. (2) Revoke the old grant ' +
  'at https://myaccount.google.com/permissions -- Google will not re-prompt ' +
  'while a partial grant is on file, so authorize() silently keeps the narrow ' +
  'token. (3) Run authorize() and tick every box on the consent screen. ' +
  '(4) Deploy > Manage deployments > edit > New version.';

/** Apps Script phrases missing-scope failures this way; a real Drive error won't. */
function isAuthError_(err) {
  var msg = (err && err.message) || String(err);
  return msg.indexOf('do not have permission') !== -1 ||
         msg.indexOf('Required permissions') !== -1 ||
         msg.indexOf('ScriptError') !== -1;
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'value']);
  }
  return sh;
}

/**
 * Bumped whenever the client needs a backend feature that older deployments
 * lack. The client reads it from getAll and warns in Settings when a stale
 * deployment is answering, because the concurrency guard below silently does
 * nothing on a script that predates it -- which is exactly the failure you
 * would not notice until two people had overwritten each other.
 */
var SERVER_VERSION = 5;

function findRow_(sh, key) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var r = 0; r < keys.length; r++) {
    if (String(keys[r][0] || '').trim() === key) return r + 2;
  }
  return 0;
}

/** One row's value, with continuation columns rejoined. */
function readRowValue_(sh, row) {
  var width = Math.max(sh.getLastColumn() - 1, 1);
  var cells = sh.getRange(row, 2, 1, width).getValues()[0];
  var value = '';
  for (var c = 0; c < cells.length; c++) {
    if (cells[c] === '' || cells[c] === null) break;
    value += String(cells[c]);
  }
  return value;
}

/** A record's field, or null when the value is not a revisioned record. */
function fieldOf_(json, name) {
  if (!json) return null;
  try {
    var o = JSON.parse(json);
    return (o && typeof o[name] === 'number') ? o[name] : null;
  } catch (err) {
    return null;   // legacy or non-record value; callers treat this as unguarded
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Every row as {key, value}, with continuation columns rejoined. */
function readAll_() {
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, Math.max(2, sh.getLastColumn())).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var key = String(rows[i][0] || '').trim();
    if (!key) continue;
    var value = '';
    for (var c = 1; c < rows[i].length; c++) {
      if (rows[i][c] === '' || rows[i][c] === null) break;
      value += String(rows[i][c]);
    }
    out.push({ key: key, value: value });
  }
  return out;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || '';

  // Does the *running deployment* actually hold Drive access? Read and write are
  // separate grants, and only write matters here: a drive.readonly token opens
  // folders happily and then fails on the first createFolder, halfway through an
  // upload. So probe both and report them apart -- canRead true with canWrite
  // false is the signature of a partial grant, which no amount of redeploying
  // fixes. The probe folder is trashed immediately.
  if (action === 'diag') {
    var drive = { canRead: false, canWrite: false, ok: false, error: null };
    try {
      drive.root = DriveApp.getRootFolder().getName();
      drive.canRead = true;
    } catch (err) {
      drive.error = err.message || String(err);
    }
    try {
      DriveApp.getRootFolder().createFolder('hermetic-diag-probe').setTrashed(true);
      drive.canWrite = true;
    } catch (err) {
      drive.error = drive.error || err.message || String(err);
    }
    drive.ok = drive.canRead && drive.canWrite;
    return json_({
      serverVersion: SERVER_VERSION,
      drive: drive,
      sheetRows: sheet_().getLastRow(),
      fix: drive.ok ? null : REAUTH_HINT
    });
  }

  // `since` turns this into a delta poll. The sheet still has to be read whole
  // either way, but the response is what actually costs -- a team polling every
  // 20s must not pull every venue's source_material each time. Rows with no
  // parseable updatedAt (legacy values, settings written by an older client)
  // are only returned on a full read, which every page load does first.
  if (action === 'getAll') {
    var prefix = p.prefix || '';
    var since = Number(p.since || 0);
    var items = readAll_().filter(function (it) {
      if (prefix && it.key.indexOf(prefix) !== 0) return false;
      if (!since) return true;
      return (fieldOf_(it.value, 'updatedAt') || 0) > since;
    });
    return json_({ items: items, serverVersion: SERVER_VERSION, now: Date.now() });
  }

  if (action === 'get') {
    var key = p.key || '';
    var hit = null;
    readAll_().forEach(function (it) { if (it.key === key) hit = it; });
    return json_({ value: hit ? hit.value : null });
  }

  if (action === 'list') {
    var pre = p.prefix || '';
    var keys = readAll_()
      .map(function (it) { return it.key; })
      .filter(function (k) { return !pre || k.indexOf(pre) === 0; });
    return json_({ keys: keys });
  }

  return json_({ error: 'Unknown action: ' + action });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ error: 'Body was not JSON: ' + err.message });
  }

  if (body.action === 'set')        return setKey_(body);
  if (body.action === 'uploadFile') return uploadFile_(body);
  if (body.action === 'copyFile')   return copyFile_(body);
  if (body.action === 'checkFiles') return checkFiles_(body);
  return json_({ error: 'Unknown action: ' + body.action });
}

function setKey_(body) {
  var key = String(body.key || '').trim();
  if (!key) return json_({ error: 'set needs a key' });
  var value = body.value == null ? '' : String(body.value);

  var chunks = chunksOf_(value);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);   // serialize writes so two saves cannot interleave
  try {
    var sh = sheet_();
    var row = findRow_(sh, key);

    // Compare-and-set. The lock alone only stops two writes from interleaving
    // mid-row; it does nothing about the second writer having loaded the record
    // before the first writer changed it. When the caller says which revision it
    // is editing and the stored row has moved on, hand back what is there now
    // and write nothing -- the client merges and retries.
    if (row && body.ifRev !== undefined && body.ifRev !== null) {
      var current = readRowValue_(sh, row);
      var curRev = fieldOf_(current, 'rev');
      // A row with no parseable rev predates revisions; let the write through
      // rather than deadlocking the client against a value it can never match.
      if (curRev !== null && curRev !== Number(body.ifRev)) {
        return json_({ conflict: true, current: current, rev: curRev });
      }
    }

    if (!row) row = Math.max(sh.getLastRow() + 1, 2);
    writeRow_(sh, row, key, chunks);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

/** A value split across continuation columns, because a cell caps at 50k. */
function chunksOf_(value) {
  var chunks = [];
  for (var i = 0; i < value.length; i += CELL_LIMIT) {
    chunks.push(value.substring(i, i + CELL_LIMIT));
  }
  return chunks.length ? chunks : [''];
}

function writeRow_(sh, row, key, chunks) {
  // Clear the old value's columns before writing, so a shorter value does not
  // leave a stale tail behind.
  var width = Math.max(sh.getLastColumn() - 1, chunks.length);
  if (width > 0) sh.getRange(row, 2, 1, width).clearContent();
  sh.getRange(row, 1).setValue(key);
  sh.getRange(row, 2, 1, chunks.length).setValues([chunks]);
}

/* ==================================================================
   EDITING THE SHEET BY HAND

   Someone fixing a typo straight in the spreadsheet is a legitimate
   way to change data, and until this trigger existed it was also a
   silent no-op: the app decides what has changed from the updatedAt
   and fieldTs stamps inside each value, and typing in a cell does not
   touch either. The delta poll therefore never returned the row, and
   the next full sync saw a field whose stamp had not moved, kept the
   copy in the browser, and wrote the hand edit back out. The edit did
   not just fail to appear -- it was reverted.

   So the trigger takes a hand edit and says it in the protocol the
   rest of the file speaks: bump the revision, move updatedAt, and
   stamp the fields that actually changed.

   onEdit is a simple trigger. It installs itself with the script and,
   importantly, does NOT fire for writes made by a script or the API,
   so setKey_ cannot set it off and it cannot recurse.
================================================================== */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEET_NAME) return;
    var first = e.range.getRow();
    if (first < 2) return;                       // header
    var last = first + e.range.getNumRows() - 1;
    for (var row = first; row <= last; row++) normalizeEditedRow_(sh, row, e);
  } catch (err) {
    // A trigger that throws fails silently for the person editing, so leave a
    // trace in the execution log rather than nothing at all.
    console.error('onEdit could not normalize the edit: ' + (err && err.message || err));
  }
}

function normalizeEditedRow_(sh, row, e) {
  var key = String(sh.getRange(row, 1).getValue() || '').trim();
  if (!key) return;
  var record;
  try {
    record = JSON.parse(readRowValue_(sh, row));
  } catch (err) {
    return;   // Hand-edited into invalid JSON. Leave it exactly as typed:
              // rewriting it would destroy whatever they were trying to say,
              // and the client already skips a row it cannot parse.
  }
  if (!record || typeof record !== 'object' || record.length !== undefined) return;

  var now = Date.now();
  record.fieldTs = record.fieldTs || {};
  var changed = changedPaths_(e, record);
  if (changed) {
    for (var i = 0; i < changed.length; i++) record.fieldTs[changed[i]] = now;
  } else {
    stampEveryField_(record, now);
  }
  record.rev = (typeof record.rev === 'number' ? record.rev : 0) + 1;
  record.updatedAt = now;
  record.lastWriter = 'sheet';
  writeRow_(sh, row, key, chunksOf_(JSON.stringify(record)));
}

/**
 * Which fields the edit actually touched, or null when that cannot be known.
 * A single-cell edit carries its previous value, so the diff is exact and only
 * the edited answer wins its merge. A paste over several cells, or an edit to a
 * continuation column, carries no usable before-value; null means "stamp
 * everything", which is blunt but never loses the edit.
 */
function changedPaths_(e, after) {
  if (!e || e.oldValue === undefined) return null;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return null;
  if (e.range.getColumn() !== 2) return null;    // a continuation column holds a fragment
  var before;
  try { before = JSON.parse(e.oldValue); } catch (err) { return null; }
  if (!before || typeof before !== 'object') return null;

  var paths = [];
  var scalars = ['name', 'drive_link', 'deleted'];
  for (var i = 0; i < scalars.length; i++) {
    var k = scalars[i];
    if (JSON.stringify(before[k]) !== JSON.stringify(after[k])) paths.push(k);
  }
  var bd = before.data || {}, ad = after.data || {}, seen = {};
  Object.keys(bd).forEach(function (k) { seen[k] = true; });
  Object.keys(ad).forEach(function (k) { seen[k] = true; });
  Object.keys(seen).forEach(function (k) {
    if (k === 'docs') return;                    // merged as a set, not by stamp
    if (JSON.stringify(bd[k]) !== JSON.stringify(ad[k])) paths.push('data.' + k);
  });
  return paths;
}

function stampEveryField_(record, now) {
  ['name', 'drive_link', 'deleted'].forEach(function (k) {
    if (record[k] !== undefined) record.fieldTs[k] = now;
  });
  var data = record.data || {};
  Object.keys(data).forEach(function (k) {
    if (k !== 'docs') record.fieldTs['data.' + k] = now;
  });
}

/**
 * Run this once by hand from the editor after any change to which Google
 * services this file touches. Apps Script decides a deployment's OAuth scopes
 * when you authorize it, not when you deploy it -- so a web app authorized
 * back when this script only read a Sheet keeps a Sheets-only token, and every
 * DriveApp call fails with "You do not have permission" no matter how many
 * times you redeploy. Running any function that touches Drive re-triggers the
 * consent screen and mints a token that covers both.
 */
function authorize() {
  var root = DriveApp.getRootFolder();
  var sh = sheet_();
  Logger.log('Authorized. Drive root: %s, kv rows: %s', root.getName(), sh.getLastRow());
}

/**
 * The destination root, preferring an explicit folder id and falling back to
 * a search by name. getFoldersByName scans all of Drive rather than just the
 * root, so a folder that lives inside another folder -- or was shared in --
 * is still found. Only when neither turns one up is a folder created.
 */
function resolveRoot_(body) {
  var id = String(body.rootFolderId || '').trim();
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (err) {
      // A stale or unshared id should not sink the upload; fall through to name.
      // A missing scope is a different animal -- every later Drive call would
      // fail the same way, so report the real cause instead of the folder that
      // could not be found.
      if (isAuthError_(err)) throw err;
    }
  }
  var name = String(body.rootFolderName || '').trim() || DEFAULT_ROOT;
  var it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.getRootFolder().createFolder(name);
}

/**
 * Saves an uploaded file into <root>/<segments...>/<category>, creating
 * folders as needed. Reconstructed from the client's uploadFile call -- if your
 * existing deployment already lays folders out differently, keep your version
 * of this function and take only the getAll branch above.
 */
/** <root>/<segments...>/<category>, creating each level that is missing. */
function destinationFolder_(body) {
  var folder = resolveRoot_(body);
  (body.segments || []).forEach(function (seg) {
    if (seg) folder = folderByName_(folder, String(seg));
  });
  if (body.category) folder = folderByName_(folder, String(body.category));
  return folder;
}

function driveResult_(file, folder) {
  return json_({
    fileId: file.getId(),
    link: file.getUrl(),
    folderId: folder.getId(),
    folderLink: folder.getUrl()
  });
}

function driveError_(err) {
  var msg = err.message || String(err);
  if (isAuthError_(err)) msg = REAUTH_HINT + ' (Google said: ' + msg + ')';
  return json_({ error: msg });
}

function uploadFile_(body) {
  try {
    var folder = destinationFolder_(body);
    var blob = Utilities.newBlob(
      Utilities.base64Decode(body.base64),
      body.mimeType || 'application/octet-stream',
      body.filename || 'upload'
    );
    return driveResult_(folder.createFile(blob), folder);
  } catch (err) {
    return driveError_(err);
  }
}

/**
 * Copies a file that is already in Drive into another venue's folder, for
 * duplicating a venue. The copy happens entirely inside Drive: sending the
 * bytes down to the browser and back up would make duplicating a venue with a
 * few large PDFs take as long as uploading them again, and would fail on
 * anything over the upload ceiling that was fine when it first went in.
 *
 * A real copy, not a shortcut to the original -- the point of duplicating a
 * venue is that the new one can have its own files edited or removed without
 * touching the venue it came from.
 */
function copyFile_(body) {
  try {
    var folder = destinationFolder_(body);
    var source = DriveApp.getFileById(String(body.fileId));
    var copy = source.makeCopy(String(body.filename || source.getName()), folder);
    return driveResult_(copy, folder);
  } catch (err) {
    return driveError_(err);
  }
}

/**
 * What actually became of each uploaded file. Drive is a shared folder that
 * people tidy: files get renamed, dragged into other folders, and thrown away,
 * none of which the console would otherwise ever hear about. It would go on
 * showing a chip that links to a file nobody can open.
 *
 * Reports per id rather than failing as a whole, because one deleted file must
 * not stop the other nine being checked. The exception is a missing Drive
 * scope: that makes every lookup throw, and reporting it per file would have
 * the console announce that every file had been deleted. That one fails loudly
 * for the whole request instead.
 */
function checkFiles_(body) {
  var ids = (body.fileIds || []).slice(0, 100);
  var out = {};
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i]);
    if (!id) continue;
    try {
      var file = DriveApp.getFileById(id);
      var parents = file.getParents();
      out[id] = {
        exists: true,
        trashed: file.isTrashed(),
        name: file.getName(),
        url: file.getUrl(),
        folderName: parents.hasNext() ? parents.next().getName() : null
      };
    } catch (err) {
      if (isAuthError_(err)) return driveError_(err);
      // getFileById throws for a file deleted outright, or one this account can
      // no longer see. Either way the console can no longer offer it.
      out[id] = { exists: false, trashed: false };
    }
  }
  return json_({ files: out });
}

function folderByName_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
