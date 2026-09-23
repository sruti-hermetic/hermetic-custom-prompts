/**
 * Hermetic AI — account setup: Apps Script backend
 *
 * Deploy: Extensions > Apps Script, paste this in. Paste appsscript.json too
 * (Project Settings > "Show appsscript.json manifest file in editor", then open
 * it from the file list). Run authorize() once from the editor and accept the
 * Drive prompt, then Deploy > New deployment > Web app, "Execute as: Me", "Who
 * has access: Anyone". Copy the /exec URL into the app's Settings > Google
 * Sheet sync. On later edits use Deploy > Manage deployments > New version,
 * which keeps the same /exec URL. Then reload the spreadsheet and run
 * Hermetic > Keep the view up to date once, and Hermetic > Check the readable
 * view to confirm it took.
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
 * THE READABLE TABS
 * kv is the protocol and is unreadable on purpose. A tab per account is built
 * beside it from the same rows, and kept up to date by one standing trigger
 * that drains a list of accounts marked stale by each save. The trigger needs
 * the script.scriptapp scope, which a deployment authorized before that scope
 * was in the manifest does not hold -- the same trap as Drive above, with the
 * same fix, and Hermetic > Check the readable view says which one you are in.
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
 * rejoined on read, so a long typed answer is never truncated.
 */

var SHEET_NAME = 'kv';
var CELL_LIMIT = 45000;   // Sheets caps a cell at 50k characters; leave headroom.
var DEFAULT_ROOT = 'ONBOARDING - Custom Prompt Uploads';

var REAUTH_STEPS =
  ' (1) Check appsscript.json lists that exact scope. (2) Revoke the old grant ' +
  'at https://myaccount.google.com/permissions -- Google will not re-prompt ' +
  'while a partial grant is on file, so authorize() silently keeps the narrow ' +
  'token. (3) Run authorize() and tick every box on the consent screen. ' +
  '(4) Deploy > Manage deployments > edit > New version.';

var REAUTH_HINT =
  'This deployment cannot write to Drive. Read-only Drive access is not enough ' +
  '-- uploads create folders and files, which needs the full ' +
  'https://www.googleapis.com/auth/drive scope.' + REAUTH_STEPS;

/** Apps Script phrases missing-scope failures this way; a real Drive error won't. */
function isAuthError_(err) {
  var msg = (err && err.message) || String(err);
  return msg.indexOf('do not have permission') !== -1 ||
         msg.indexOf('Required permissions') !== -1 ||
         msg.indexOf('ScriptError') !== -1;
}

/**
 * Different calls need different scopes -- DriveApp needs drive, the direct
 * UrlFetchApp.fetch in startUpload_ needs script.external_request -- and
 * Apps Script names the missing one right in the error ("Required
 * permissions: <scope>"). Naming it beats REAUTH_HINT's Drive-only wording,
 * which sends someone to check a scope that was never the problem.
 */
function scopeErrorHint_(err) {
  var msg = (err && err.message) || String(err);
  var m = msg.match(/Required permissions?:\s*(\S+)/);
  if (!m) return REAUTH_HINT;
  return 'This deployment is missing a scope it now needs: ' + m[1] + '.' + REAUTH_STEPS;
}

/**
 * Errors DriveApp throws that say nothing about whether a file still exists --
 * a per-user rate limit tripped by a team all checking files at once, or a
 * transient hiccup in Drive itself. checkFiles_ used to treat every non-auth
 * error the same as "this id is gone", so a rate limit hit mid-batch reported
 * every file after it as deleted from Drive -- including one someone had
 * uploaded a minute earlier. Recognized separately so those ids are left
 * unreported instead of wrongly marked gone.
 */
function isTransientDriveError_(err) {
  var msg = (err && err.message) || String(err);
  return msg.indexOf('too many times') !== -1 ||      // rate limit
         msg.indexOf('Internal error') !== -1 ||
         msg.indexOf('Service unavailable') !== -1 ||
         msg.indexOf('try again') !== -1;
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
var SERVER_VERSION = 8;

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

/**
 * Every row as {key, value}, with continuation columns rejoined, briefly
 * cached. getAll is a bulk read of the whole sheet -- every client polls it
 * every 20s, on top of every page load -- and a team's kv grows without
 * bound, so the read it pays for keeps getting bigger. Apps Script also runs
 * one request at a time per user, so a slow read here queues up whoever is
 * behind it, which is exactly the "someone's new account takes a while to
 * show up" complaint: not a wrong answer, a slow one, at the worst possible
 * layer to be slow at.
 *
 * A short cache lets several polls landing within the same few seconds --
 * normal with more than one or two people on the tool -- share one sheet
 * read instead of paying for their own. KV_CACHE_TTL_S bounds how stale that
 * shared answer can be; invalidateReadCache_ (called right after every
 * write) keeps the common case -- no write racing a read -- serving fresh
 * data immediately, and the TTL is only what a genuinely concurrent write
 * can hide behind.
 */
var KV_CACHE_KEY = 'hermetic.kvCache';
var KV_CACHE_TTL_S = 5;

function readAll_() {
  var cache = kvCache_();
  if (cache) {
    try {
      var cached = cache.get(KV_CACHE_KEY);
      if (cached) return JSON.parse(cached);
    } catch (err) {}   // an unreadable entry is no worse than no entry
  }

  var sh = sheet_();
  var last = sh.getLastRow();
  var out = [];
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, Math.max(2, sh.getLastColumn())).getValues();
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
  }

  if (cache) {
    try { cache.put(KV_CACHE_KEY, JSON.stringify(out), KV_CACHE_TTL_S); }
    catch (err) {}    // over the 100KB cache entry limit; every call just reads the sheet
  }
  return out;
}

function kvCache_() {
  try { return CacheService.getDocumentCache(); } catch (err) { return null; }
}

/** Called right after any write lands, so the next read is never the one
 *  serving a just-written row out of a cache built before it existed. */
function invalidateReadCache_() {
  var cache = kvCache_();
  if (cache) { try { cache.remove(KV_CACHE_KEY); } catch (err) {} }
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
    // Row counts by kind: a venue with no account row of its own is the shape
    // of a spreadsheet whose tabs cannot be built, and it is invisible from the
    // console, which only ever shows accounts it can already see.
    var rows = { accounts: 0, venues: 0, schema: false, orphanVenues: 0 };
    var accountIds = {}, venueOwners = {};
    readAll_().forEach(function (it) {
      if (it.key === 'hermetic:schema') { rows.schema = true; return; }
      if (it.key.indexOf('hermetic:account:') === 0) {
        rows.accounts++; accountIds[it.key.substring(17)] = true; return;
      }
      if (it.key.indexOf('hermetic:venue:') === 0) {
        rows.venues++; venueOwners[it.key.split(':')[2]] = true;
      }
    });
    Object.keys(venueOwners).forEach(function (id) { if (!accountIds[id]) rows.orphanVenues++; });

    return json_({
      serverVersion: SERVER_VERSION,
      drive: drive,
      sheetRows: sheet_().getLastRow(),
      rows: rows,
      view: viewStatus_(),
      fix: drive.ok ? null : REAUTH_HINT
    });
  }

  // `since` turns this into a delta poll. The sheet still has to be read whole
  // either way, but the response is what actually costs -- a team polling every
  // 20s must not pull every venue's answers each time. Rows with no
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

  if (body.action === 'set')         return setKey_(body);
  if (body.action === 'startUpload') return startUpload_(body);
  if (body.action === 'uploadFile')  return uploadFile_(body);
  if (body.action === 'copyFile')    return copyFile_(body);
  if (body.action === 'checkFiles')  return checkFiles_(body);
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
    invalidateReadCache_();
  } finally {
    lock.releaseLock();
  }

  // The readable tabs are downstream of this write, so mark them stale --
  // outside the lock, and never in a way that can fail the save itself.
  markViewDirty_(key);
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
  invalidateReadCache_();
  // Typing straight into kv is a save like any other, so the tab that displays
  // it has to follow. Without this the only edits the view ever saw were the
  // ones that came through the console.
  markViewDirty_(key);
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
  // Touch ScriptApp too: the refresh after a save creates a one-off trigger,
  // and a deployment authorized before that existed holds a token without the
  // scope for it -- so the save would work and the readable tabs would not.
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Authorized. Drive root: %s, kv rows: %s, triggers: %s',
             root.getName(), sh.getLastRow(), triggers.length);
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
 * The venue's own folder -- <root>/<segments...> -- for listing what is
 * really in it, never for writing. Unlike resolveRoot_/destinationFolder_,
 * nothing here is created: a folder that does not exist yet just means
 * nothing has been uploaded for this venue, which is not an error, only an
 * empty answer. Returns null the moment any level of the path is missing.
 */
function findFolder_(body) {
  var id = String(body.rootFolderId || '').trim();
  var folder = null;
  if (id) {
    try { folder = DriveApp.getFolderById(id); }
    catch (err) { if (isAuthError_(err)) throw err; }
  }
  if (!folder) {
    var name = String(body.rootFolderName || '').trim() || DEFAULT_ROOT;
    var rootIt = DriveApp.getFoldersByName(name);
    if (!rootIt.hasNext()) return null;
    folder = rootIt.next();
  }
  var segments = body.segments || [];
  for (var i = 0; i < segments.length; i++) {
    if (!segments[i]) continue;
    var it = folder.getFoldersByName(String(segments[i]));
    if (!it.hasNext()) return null;
    folder = it.next();
  }
  return folder;
}

/**
 * Every non-trashed file under a folder, one level of category subfolders
 * deep -- a file filed under a category is labelled with that category's
 * name; one sitting loose in the venue folder itself (or a level the console
 * never made, from someone tidying by hand) is labelled with the label the
 * caller is already using at that depth. Capped in both directions so a
 * folder somebody has turned into a general dumping ground cannot make one
 * request read all of Drive.
 */
var LIST_FILES_DEPTH = 2;
var LIST_FILES_MAX = 300;
function collectFiles_(folder, label, out, depth) {
  if (depth > LIST_FILES_DEPTH || out.length >= LIST_FILES_MAX) return;
  var files = folder.getFiles();
  while (files.hasNext() && out.length < LIST_FILES_MAX) {
    var f = files.next();
    if (f.isTrashed()) continue;
    out.push({ fileId: f.getId(), name: f.getName(), category: label, link: f.getUrl() });
  }
  var subs = folder.getFolders();
  while (subs.hasNext() && out.length < LIST_FILES_MAX) {
    var sub = subs.next();
    collectFiles_(sub, sub.getName(), out, depth + 1);
  }
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
  if (isAuthError_(err)) msg = scopeErrorHint_(err) + ' (Google said: ' + msg + ')';
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
 * Opens a Drive resumable-upload session and hands the URL straight to the
 * browser, which then PUTs the file bytes directly to Drive. The file never
 * passes through doPost: a base64 JSON body here would inflate the file by
 * ~33% and run into the size ceiling Apps Script web apps put on request
 * bodies, which is what used to make anything much over a few MB fail (see
 * the client's putFileToUploadUrl for the other half of this).
 *
 * ScriptApp.getOAuthToken() works for an anonymous visitor because this web
 * app is deployed to execute as the deploying user (see appsscript.json) --
 * every Drive call in this file, including the ones above, already runs as
 * that same account.
 */
function startUpload_(body) {
  try {
    var folder = destinationFolder_(body);
    var metadata = {
      name: body.filename || 'upload',
      parents: [folder.getId()]
    };
    if (body.mimeType) metadata.mimeType = body.mimeType;
    var initRes = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,webViewLink',
      {
        method: 'post',
        contentType: 'application/json; charset=UTF-8',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify(metadata),
        muteHttpExceptions: true
      }
    );
    if (initRes.getResponseCode() >= 300) {
      return json_({ error: 'Drive would not start the upload: ' + initRes.getContentText() });
    }
    var headers = initRes.getHeaders();
    var uploadUrl = headers['Location'] || headers['location'];
    if (!uploadUrl) return json_({ error: 'Drive did not return an upload session.' });
    return json_({ uploadUrl: uploadUrl, folderId: folder.getId(), folderLink: folder.getUrl() });
  } catch (err) {
    return driveError_(err);
  }
}

/**
 * Copies a file that is already in Drive into another venue's folder, for
 * duplicating a venue. The copy happens entirely inside Drive: sending the
 * bytes down to the browser and back up would make duplicating a venue with a
 * few large PDFs take as long as uploading them again.
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
 * What actually became of each uploaded file, and -- when the caller sends
 * `segments` -- what Drive holds for this venue that the console does not
 * know about at all yet. Drive is a shared folder that people tidy: files
 * get renamed, dragged into other folders, thrown away, or dropped straight
 * in by hand rather than through the console's own upload button, none of
 * which the console would otherwise ever hear about. Without the second half
 * it would go on showing a chip that links to a file nobody can open, and
 * would never notice a menu that landed in the folder any way other than its
 * own Upload button.
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
      // A rate limit or a Drive hiccup says nothing about this file; reporting
      // it as gone would be a guess, and the console takes silence as "no news"
      // and leaves the file exactly as it was.
      if (isTransientDriveError_(err)) continue;
      // Anything else is getFileById throwing because the file was deleted
      // outright, or this account can no longer see it. Either way the console
      // can no longer offer it.
      out[id] = { exists: false, trashed: false };
    }
  }

  var result = { files: out };
  if (body.segments) {
    try {
      var folder = findFolder_(body);
      var found = [];
      if (folder) collectFiles_(folder, 'Other', found, 0);
      result.found = found;
    } catch (err) {
      if (isAuthError_(err)) return driveError_(err);
      // Could not list the folder for some other reason (a rate limit, a
      // hiccup): say nothing about what is in it rather than reporting it
      // empty, which the console would read as "every file in Drive is one I
      // already know about" and never ask again this session.
    }
  }
  return json_(result);
}

/* ==================================================================
   THE READABLE VIEW

   The kv tab is a protocol, not a report. Every row is one JSON
   record because that is what makes concurrent editing safe: a
   revision to compare, a stamp per field to merge on. It is also
   unreadable, and the people who live in this spreadsheet are not the
   people who wrote it.

   So rather than change kv, this builds a tab per account beside it:
   one row per section, one column per location, with the times things
   were created and last touched at the top. Generated, disposable,
   and rebuilt from kv -- kv stays the single source of truth and
   nothing here can drift from it.

   The section names come from the console, published into a
   hermetic:schema row, so renaming a section there renames the row
   here without anyone maintaining a second list.

   HOW A SAVE REACHES THE TAB
   A save marks the account it touched dirty -- one document property,
   written outside the lock, which cannot fail the save -- and a single
   standing trigger drains the dirty list once a minute. Only the
   accounts on that list are rebuilt, so one specialist typing into one
   venue costs one tab, not the whole spreadsheet.

   This replaced a scheme that created a one-off trigger per save.
   That had two failure modes and hit both: a project is capped at 20
   triggers, and a trigger that fails to fire is never cleaned up, so
   the caps filled and every later attempt to queue a refresh threw --
   silently, by design, because a refresh must never cost somebody
   their answer. The view then stopped moving for good. One standing
   trigger cannot accumulate, and ensureViewTrigger_ puts it back if it
   is ever missing.
================================================================== */
var VIEW_REGISTRY_KEY = 'hermetic.viewTabs';      // {accountId: tab name}
var VIEW_DIRTY_PREFIX = 'hermetic.dirty.';        // one property per account awaiting a rebuild
var VIEW_DIRTY_ALL    = '*';                      // stands for "every account"
var VIEW_STATUS_KEY   = 'hermetic.viewStatus';    // what the last drain did, for ?action=diag
var VIEW_SWEPT_KEY    = 'hermetic.viewSweptAt';
var VIEW_CHECKED_KEY  = 'hermetic.viewTriggerCheckedAt';
var VIEW_TICK         = 'drainViewQueue';         // the one standing trigger's handler
var VIEW_SWEEP_MS     = 6 * 60 * 60 * 1000;       // belt and braces: a full rebuild this often
var VIEW_BUDGET_MS    = 4 * 60 * 1000;            // stop well inside the 6 minute wall
var VIEW_RECHECK_MS   = 10 * 60 * 1000;           // how often a save re-checks the trigger exists
var VIEW_CELL_LIMIT   = 2000;   // keep one long transcript from swallowing a row
var VIEW_NOTE = 'Generated from the console. Anything typed on this tab is replaced the next time it refreshes, so make changes in the tool.';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Hermetic')
    .addItem('Refresh readable view now', 'rebuildViews')
    .addItem('Check the readable view', 'showViewStatus')
    .addSeparator()
    .addItem('Show the raw data tab', 'showRawTab')
    .addItem('Hide the raw data tab', 'hideRawTab')
    .addSeparator()
    .addItem('Keep the view up to date', 'installViewRefresh')
    .addItem('Stop keeping it up to date', 'removeViewRefresh')
    .addToUi();
}

/* ---- The standing trigger ------------------------------------------------ */

function viewTriggers_() {
  // rebuildViews and rebuildViewsQueued are the handlers older versions of this
  // file installed. They are collected here so switching over removes them.
  var handlers = { drainViewQueue: 1, rebuildViews: 1, rebuildViewsQueued: 1 };
  return ScriptApp.getProjectTriggers().filter(function (t) {
    return handlers[t.getHandlerFunction()] === 1;
  });
}

function installViewRefresh() {
  removeViewRefresh();
  ScriptApp.newTrigger(VIEW_TICK).timeBased().everyMinutes(1).create();
  PropertiesService.getDocumentProperties().setProperty(VIEW_CHECKED_KEY, String(Date.now()));
  markEverythingDirty_();
  drainViewQueue();
  SpreadsheetApp.getActive().toast(
    'The readable view will now follow the tool, about a minute behind.', 'Hermetic', 6);
}

function removeViewRefresh() {
  viewTriggers_().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  PropertiesService.getDocumentProperties().deleteProperty(VIEW_CHECKED_KEY);
}

/**
 * Put the standing trigger back if it has gone -- someone ran "stop", a
 * migration removed it, the project was copied. Called on every save, so it
 * only actually looks every VIEW_RECHECK_MS: getProjectTriggers is a round trip
 * and a save is on somebody's critical path.
 */
function ensureViewTrigger_(force) {
  var props = PropertiesService.getDocumentProperties();
  var checked = Number(props.getProperty(VIEW_CHECKED_KEY) || 0);
  if (!force && checked && Date.now() - checked < VIEW_RECHECK_MS) return;
  var live = viewTriggers_();
  var good = live.filter(function (t) { return t.getHandlerFunction() === VIEW_TICK; });
  // Anything that is not our one minute trigger is a leftover taking up a slot.
  live.forEach(function (t) {
    if (t.getHandlerFunction() !== VIEW_TICK || t !== good[0]) ScriptApp.deleteTrigger(t);
  });
  if (!good.length) ScriptApp.newTrigger(VIEW_TICK).timeBased().everyMinutes(1).create();
  props.setProperty(VIEW_CHECKED_KEY, String(Date.now()));
}

/* ---- The dirty list ------------------------------------------------------ */

/** Which account a key belongs to: an id, '*' for all of them, or null. */
function accountOfKey_(key) {
  if (key === 'hermetic:schema') return VIEW_DIRTY_ALL;   // section titles, on every tab
  if (key.indexOf('hermetic:account:') === 0) return key.substring(17);
  if (key.indexOf('hermetic:venue:') === 0) return key.split(':')[2] || null;
  return null;   // presence and settings change nothing anyone reads here
}

/**
 * Note that an account needs rebuilding. One property per account rather than
 * one list, because two saves landing together would each read the list, add
 * their own id, and write it back -- and the second would lose the first.
 * Distinct keys cannot collide that way.
 */
function markViewDirty_(key) {
  try {
    var account = accountOfKey_(key);
    if (!account) return;
    PropertiesService.getDocumentProperties()
      .setProperty(VIEW_DIRTY_PREFIX + account, String(Date.now()));
    ensureViewTrigger_(false);
  } catch (err) {
    // Saving the customer's answer matters more than the tab that displays it.
    // A lost mark costs a stale tab until the next save or the next sweep;
    // throwing here would cost the answer.
    console.error('Could not queue a view refresh: ' + (err && err.message || err));
  }
}

function markEverythingDirty_() {
  PropertiesService.getDocumentProperties()
    .setProperty(VIEW_DIRTY_PREFIX + VIEW_DIRTY_ALL, String(Date.now()));
}

/** The dirty marks, as {accountId: stamp}. */
function readDirty_(props) {
  var all = props.getProperties();
  var out = {};
  Object.keys(all).forEach(function (k) {
    if (k.indexOf(VIEW_DIRTY_PREFIX) === 0) out[k.substring(VIEW_DIRTY_PREFIX.length)] = all[k];
  });
  return out;
}

/**
 * Clear a mark, but only if it still says what it said when we read it. A save
 * that landed while this account was being rebuilt has moved the stamp on, and
 * deleting it would drop that edit until the next sweep.
 */
function clearDirty_(props, account, stamp) {
  if (props.getProperty(VIEW_DIRTY_PREFIX + account) !== stamp) return;
  props.deleteProperty(VIEW_DIRTY_PREFIX + account);
}

/* ---- The drain ----------------------------------------------------------- */

/**
 * The standing trigger's handler, and the only thing that writes a view tab.
 * Rebuilds the accounts on the dirty list and nothing else, within a time
 * budget: whatever it does not reach stays marked and is picked up a minute
 * later, so a spreadsheet with more accounts than fit in one execution still
 * converges instead of dying on the wall clock.
 */
function drainViewQueue() {
  // A rebuild can outlast the minute between ticks. Two of them interleaving
  // would fight over the same tabs, so the second simply stands down -- its
  // work is still on the dirty list and the next tick will take it.
  var lock = LockService.getDocumentLock();
  try { if (!lock.tryLock(1000)) return; } catch (err) { return; }

  var started = Date.now();
  var props = PropertiesService.getDocumentProperties();
  var status = { at: started, built: [], dropped: [], left: 0, ms: 0, error: null };
  try {
    var dirty = readDirty_(props);
    var sweptAt = Number(props.getProperty(VIEW_SWEPT_KEY) || 0);
    var sweepDue = !sweptAt || started - sweptAt > VIEW_SWEEP_MS;
    if (!Object.keys(dirty).length && !sweepDue) return;   // nothing to do: the common case

    var model = readModelSafely_();
    var registry = readRegistry_();

    // "Everything" is expanded into a mark per account before any work starts,
    // so a sweep that runs out of budget resumes where it stopped rather than
    // starting over. Registry ids are included: an account whose row has gone
    // is exactly the one whose tab needs removing.
    if (dirty[VIEW_DIRTY_ALL] || sweepDue) {
      var marks = {};
      var stamp = String(started);
      model.accounts.forEach(function (a) { marks[VIEW_DIRTY_PREFIX + a.id] = stamp; });
      Object.keys(registry).forEach(function (id) { marks[VIEW_DIRTY_PREFIX + id] = stamp; });
      if (Object.keys(marks).length) props.setProperties(marks);
      props.deleteProperty(VIEW_DIRTY_PREFIX + VIEW_DIRTY_ALL);
      props.setProperty(VIEW_SWEPT_KEY, stamp);
      dirty = readDirty_(props);
      delete dirty[VIEW_DIRTY_ALL];
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var byId = {};
    model.accounts.forEach(function (a) { byId[a.id] = a; });

    var deadline = started + VIEW_BUDGET_MS;
    var ids = Object.keys(dirty).sort();
    var structural = false;

    for (var i = 0; i < ids.length; i++) {
      // The budget is checked after the first account and never before it. A run
      // that begins already over -- a slow read of a large kv, a tick that waited
      // on the lock -- would otherwise build nothing, and so would every tick
      // after it: the view would stop for good while the list only grew. One
      // account per run is slow; no accounts per run is broken.
      if (i > 0 && Date.now() > deadline) { status.left = ids.length - i; break; }
      var id = ids[i];
      if (byId[id]) {
        var before = registry[id];
        registry[id] = buildAccountView_(ss, byId[id], model, registry);
        if (registry[id] !== before) structural = true;
        status.built.push(registry[id]);
      } else if (registry[id]) {
        // The account was deleted, or its row never existed. Either way the tab
        // it used to own has nothing behind it now.
        var stale = ss.getSheetByName(registry[id]);
        if (stale && ss.getSheets().length > 1) ss.deleteSheet(stale);
        status.dropped.push(registry[id]);
        delete registry[id];
        structural = true;
      }
      clearDirty_(props, id, dirty[id]);
      writeRegistry_(registry);
    }

    var names = Object.keys(registry).map(function (k) { return registry[k]; });
    // Nothing readable left to show -- the last account went, or none has been
    // made yet. Bring kv back before the tabs go, because Sheets will not let
    // the file end up with every tab hidden.
    if (!names.length) {
      var raw = ss.getSheetByName(SHEET_NAME);
      if (raw && raw.isSheetHidden()) raw.showSheet();
    } else if (structural) {
      tidyTabs_(ss, names);
    }
    SpreadsheetApp.flush();
  } catch (err) {
    status.error = (err && err.message) || String(err);
    console.error('Readable view rebuild failed: ' + status.error);
  } finally {
    status.ms = Date.now() - started;
    if (status.built.length > 12) status.built = status.built.slice(0, 12).concat(['+' + (status.built.length - 12) + ' more']);
    try { props.setProperty(VIEW_STATUS_KEY, JSON.stringify(status)); } catch (e) {}
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** The menu's "refresh now": mark the lot and drain it here and now. */
function rebuildViews() {
  markEverythingDirty_();
  PropertiesService.getDocumentProperties().deleteProperty(VIEW_SWEPT_KEY);
  drainViewQueue();
  var left = Object.keys(readDirty_(PropertiesService.getDocumentProperties())).length;
  try {
    SpreadsheetApp.getActive().toast(
      left ? left + ' more to go; the rest follow within a minute or two.' : 'The readable view is up to date.',
      'Hermetic', 5);
  } catch (err) {}
  return left;
}

/** A trigger left over from the previous scheme. Hand the work on and go. */
function rebuildViewsQueued() {
  try { ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rebuildViewsQueued') ScriptApp.deleteTrigger(t);
  }); } catch (err) {}
  drainViewQueue();
}

/* ---- What the view knows ------------------------------------------------- */

function readRegistry_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(VIEW_REGISTRY_KEY);
  var parsed;
  try { parsed = JSON.parse(raw || '{}'); } catch (err) { return {}; }
  // The previous scheme stored a bare list of tab names with no account ids.
  // There is nothing to recover from it; the tabs are re-adopted by name below.
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return {};
  return parsed;
}

function writeRegistry_(registry) {
  PropertiesService.getDocumentProperties().setProperty(VIEW_REGISTRY_KEY, JSON.stringify(registry));
}

/** Everything in kv, parsed into accounts and their locations. */
function readModel_() {
  var model = { accounts: [], venues: {}, schema: null };
  var seenAccount = {};
  readAll_().forEach(function (r) {
    var o;
    try { o = JSON.parse(r.value); } catch (err) { return; }
    if (!o || typeof o !== 'object') return;

    if (r.key === 'hermetic:schema') { model.schema = o; return; }
    if (r.key.indexOf('hermetic:account:') === 0) {
      seenAccount[r.key.substring(17)] = true;
      if (!o.deleted) model.accounts.push(o);
      return;
    }
    if (r.key.indexOf('hermetic:venue:') === 0) {
      if (o.deleted) return;
      var accountId = r.key.split(':')[2];
      (model.venues[accountId] = model.venues[accountId] || []).push(o);
    }
  });

  // Locations whose account row is missing entirely -- never written, or lost.
  // Without this they are invisible here: the view is built from account rows,
  // so a whole customer's answers can sit in kv with no tab to read them on.
  // An account row that exists and says deleted is not this case and stays gone.
  Object.keys(model.venues).forEach(function (id) {
    if (seenAccount[id]) return;
    var created = model.venues[id].reduce(function (min, v) {
      return Math.min(min, v.createdAt || Date.now());
    }, Date.now());
    model.accounts.push({ id: id, name: 'Unfiled locations', createdAt: created, orphan: true });
  });

  var byCreated = function (a, b) { return (a.createdAt || 0) - (b.createdAt || 0); };
  model.accounts.sort(byCreated);
  Object.keys(model.venues).forEach(function (k) { model.venues[k].sort(byCreated); });
  return model;
}

/**
 * kv, read under the lock when the lock is available and without it when it is
 * not. A row caught mid-write is half cleared and fails to parse, so readModel_
 * skips it and that one location is missing until the next refresh. A view that
 * is briefly one location short beats a view that refuses to build at all
 * because saves are busy.
 */
function readModelSafely_() {
  var lock = LockService.getScriptLock();
  var held = false;
  try { held = lock.tryLock(10000); } catch (err) { held = false; }
  try {
    return readModel_();
  } finally {
    if (held) lock.releaseLock();
  }
}

/**
 * What the spreadsheet opens on. The readable tabs go first, and kv -- plus the
 * empty Sheet1 Google creates with every spreadsheet -- get hidden behind them,
 * because the first thing anyone saw on opening this file was a wall of JSON.
 * Hidden, not deleted: kv is still the only source of truth, the script still
 * reads and writes it by name, and the Hermetic menu brings it back in a click.
 *
 * Only run when a tab has appeared or gone. Reordering every tab on every save
 * is a pile of Sheets calls to reach the order they were already in, and it
 * yanks the active tab out from under anyone reading the spreadsheet.
 */
function tidyTabs_(ss, built) {
  var order = built.slice().sort();
  for (var i = 0; i < order.length; i++) {
    var sh = ss.getSheetByName(order[i]);
    if (!sh) continue;
    if (sh.isSheetHidden()) sh.showSheet();
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  }
  var first = ss.getSheetByName(order[0]);
  if (first) ss.setActiveSheet(first);

  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (order.indexOf(name) !== -1 || sh.isSheetHidden()) return;
    if (name === SHEET_NAME) { sh.hideSheet(); return; }
    // Only the untouched default tab, never a sheet somebody has put work in.
    if (/^Sheet\d+$/.test(name) && sh.getLastRow() === 0 && sh.getLastColumn() === 0) sh.hideSheet();
  });
}

function showRawTab() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return;
  sh.showSheet();
  SpreadsheetApp.getActive().toast(
    'The ' + SHEET_NAME + ' tab is the raw data the tool reads and writes. Editing it by hand works, but the tool is safer.',
    'Hermetic', 8);
}

function hideRawTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) return;
  var visible = ss.getSheets().filter(function (s) { return !s.isSheetHidden(); });
  if (visible.length < 2) {
    ss.toast('There is no other tab to show yet. Refresh the readable view first.', 'Hermetic', 6);
    return;
  }
  if (ss.getActiveSheet().getName() === SHEET_NAME) {
    for (var i = 0; i < visible.length; i++) {
      if (visible[i].getName() !== SHEET_NAME) { ss.setActiveSheet(visible[i]); break; }
    }
  }
  sh.hideSheet();
}

/**
 * Why the view is or is not moving, in one place. The whole machine is
 * invisible when it works and equally invisible when it does not -- a save that
 * cannot queue a refresh is swallowed on purpose -- so this is the only way to
 * tell the two apart without reading the execution log. Also served from
 * ?action=diag, so it can be checked without opening the spreadsheet.
 */
function viewStatus_() {
  var props = PropertiesService.getDocumentProperties();
  var status = {};
  try { status = JSON.parse(props.getProperty(VIEW_STATUS_KEY) || '{}'); } catch (err) {}
  var out = {
    trigger: false,
    triggerCount: 0,
    pending: Object.keys(readDirty_(props)),
    lastRunAt: status.at || null,
    lastRunMs: status.ms || null,
    lastBuilt: status.built || [],
    lastError: status.error || null,
    sweptAt: Number(props.getProperty(VIEW_SWEPT_KEY) || 0) || null,
    tabs: readRegistry_(),
    triggerError: null
  };
  try {
    var live = viewTriggers_();
    out.triggerCount = live.length;
    out.trigger = live.some(function (t) { return t.getHandlerFunction() === VIEW_TICK; });
    out.projectTriggers = ScriptApp.getProjectTriggers().length;
  } catch (err) {
    // The signature of a deployment authorized before script.scriptapp was in
    // the manifest: it can write the sheet and cannot own a trigger, so the
    // view never refreshes itself and nothing says why.
    out.triggerError = (err && err.message) || String(err);
  }
  return out;
}

function showViewStatus() {
  var v = viewStatus_();
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var lines = [];
  lines.push(v.trigger
    ? 'Automatic refresh: ON (every minute).'
    : 'Automatic refresh: OFF. Run Hermetic > Keep the view up to date.');
  if (v.triggerError) {
    lines.push('');
    lines.push('This script cannot own a trigger: ' + v.triggerError);
    lines.push('Open Extensions > Apps Script, run authorize(), accept every prompt, ' +
               'then Deploy > Manage deployments > edit > New version.');
  }
  lines.push('Accounts with a tab: ' + Object.keys(v.tabs).length);
  lines.push('Waiting to be rebuilt: ' + v.pending.length);
  lines.push('Last checked: ' + (v.lastRunAt ? when_(v.lastRunAt, tz) + ' (' +
             Math.round((v.lastRunMs || 0) / 100) / 10 + 's)' : 'never'));
  if (v.lastBuilt.length) lines.push('Last rebuilt: ' + v.lastBuilt.join(', '));
  if (v.lastError) lines.push('Last error: ' + v.lastError);
  SpreadsheetApp.getUi().alert('Hermetic readable view', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ---- Building one tab ---------------------------------------------------- */

/** A tab name Sheets will accept, and that is not already spoken for. */
function safeSheetName_(name, taken) {
  var base = String(name || 'Account').replace(/[\[\]\*\/\\\?:]/g, ' ').trim().substring(0, 90) || 'Account';
  if (base.toLowerCase() === SHEET_NAME) base = base + ' (account)';
  var candidate = base, n = 2;
  while (taken.indexOf(candidate) !== -1) candidate = base + ' ' + (n++);
  return candidate;
}

function when_(ms, tz) {
  if (!ms) return '';
  return Utilities.formatDate(new Date(Number(ms)), tz, 'd MMM yyyy, h:mm a');
}

function cellText_(value) {
  if (value === null || value === undefined) return '';
  var s = String(value).trim();
  if (s.length <= VIEW_CELL_LIMIT) return s;
  return s.substring(0, VIEW_CELL_LIMIT) + '\n\n[Shortened for this view. The whole answer is in the tool.]';
}

/* The section list the console published. Falling back to the keys the data
   happens to carry keeps the view useful on a spreadsheet whose console has
   not loaded since this was added, rather than showing nothing at all. */
function viewParts_(model, locations) {
  if (model.schema && model.schema.parts && model.schema.parts.length) return model.schema.parts;
  // Not answers: files and folder links, plus the retired source-material and
  // generated-output boxes, whose values still sit in records typed before the
  // console dropped them.
  var skip = { docs: 1, drive_link: 1, source_material: 1, draft_location_doc: 1, draft_campaign_doc: 1, draft_open_questions: 1 };
  var keys = {};
  locations.forEach(function (l) {
    Object.keys(l.data || {}).forEach(function (k) { if (!skip[k]) keys[k] = true; });
  });
  return [{ title: 'Answers', sections: Object.keys(keys).sort().map(function (k) {
    var title = k.replace(/_/g, ' ').replace(/ text$/, '');
    return { key: k, title: title.charAt(0).toUpperCase() + title.slice(1) };
  }) }];
}

function fileSummary_(location) {
  var docs = (location.data || {}).docs || [];
  if (!docs.length) return '';
  return docs.map(function (d) {
    var state = d.gone === 'trashed' ? '  [in Drive trash]' : (d.gone === 'deleted' ? '  [deleted from Drive]' : '');
    return '- ' + (d.category || 'Other') + ': ' + (d.name || 'file') + state;
  }).join('\n');
}

/**
 * The tab this account owns, adopting the one it had last time even if the
 * account has since been renamed. Names are claimed by account id, so two
 * accounts called the same thing get a tab each -- before this, the second one
 * found the first one's tab by name and overwrote it, and one customer's
 * answers silently replaced another's.
 */
function accountSheet_(ss, account, registry) {
  var claimed = {};
  Object.keys(registry).forEach(function (id) {
    if (id !== account.id) claimed[registry[id]] = true;
  });
  var taken = ss.getSheets().map(function (s) { return s.getName(); })
    .filter(function (n) { return claimed[n] || n === SHEET_NAME; });

  var wanted = safeSheetName_(account.name, taken);
  var mine = registry[account.id] ? ss.getSheetByName(registry[account.id]) : null;
  if (mine) {
    if (mine.getName() !== wanted && !ss.getSheetByName(wanted)) mine.setName(wanted);
    return mine;
  }
  // No tab on record. Adopt one standing under this name -- that is how tabs
  // built before the registry existed are picked back up instead of duplicated.
  var byName = ss.getSheetByName(wanted);
  if (byName && !claimed[wanted] && wanted !== SHEET_NAME) return byName;
  return ss.insertSheet(safeSheetName_(account.name,
    ss.getSheets().map(function (s) { return s.getName(); })));
}

function buildAccountView_(ss, account, model, registry) {
  var tz = ss.getSpreadsheetTimeZone();
  var locations = model.venues[account.id] || [];
  var sh = accountSheet_(ss, account, registry);
  // Last time's merged note row spans the columns it had then. An account that
  // has since lost a location writes a narrower grid, and setValues over a
  // merge it only half covers throws -- so break the merges before clearing.
  if (sh.getLastRow()) sh.getRange(1, 1, sh.getLastRow(), sh.getMaxColumns()).breakApart();
  sh.clear();   // contents and formats both; clearFormats afterwards is a second round trip for nothing

  var header = [''].concat(locations.map(function (l) { return l.name || 'Untitled'; }));
  var grid = [];
  var groupRows = [];      // rows to style as a section heading
  var blank = locations.map(function () { return ''; });
  var push = function (row) { grid.push(row); return grid.length; };
  var group = function (title) { groupRows.push(push([title].concat(blank))); };
  var line = function (label, valueOf) { push([label].concat(locations.map(valueOf))); };

  push([account.name || 'Account'].concat(blank));
  push([(account.orphan
    ? 'These locations have no account record in the tool. They are shown here so the answers are not lost. '
    : '') + VIEW_NOTE].concat(blank));
  push(['Refreshed ' + when_(Date.now(), tz)].concat(blank));
  push([''].concat(blank));
  var headerRow = push(header);

  group('WHEN');
  line('Location created', function (l) { return when_(l.createdAt, tz); });
  line('Answers last edited', function (l) { return when_(l.lastEditedAt, tz); });
  line('Last edited by', function (l) { return l.lastWriterName || ''; });
  line('Last saved to this sheet', function (l) { return when_(l.updatedAt, tz); });
  // Edited after this, and the console shows the location as outdated: the
  // answers have moved on since anybody last handed them to Claude.
  line('Last copied for Claude', function (l) { return when_(l.lastCopiedAt, tz); });

  group('FILES');
  line('Uploaded files', function (l) { return fileSummary_(l); });
  line('Drive folder', function (l) { return (l.data || {}).drive_link || ''; });

  viewParts_(model, locations).forEach(function (part) {
    group(String(part.title || 'Answers').toUpperCase());
    (part.sections || []).forEach(function (section) {
      line(section.title, function (l) { return cellText_((l.data || {})[section.key]); });
    });
  });

  if (!locations.length) push(['This account has no locations yet.']);

  var width = Math.max(header.length, 1);
  grid = grid.map(function (row) {
    while (row.length < width) row.push('');
    return row;
  });
  sh.getRange(1, 1, grid.length, width).setValues(grid);
  styleAccountView_(sh, grid.length, width, headerRow, groupRows);
  return sh.getName();
}

function styleAccountView_(sh, rows, width, headerRow, groupRows) {
  var all = sh.getRange(1, 1, rows, width);
  all.setVerticalAlignment('top').setFontFamily('Arial').setFontSize(10).setWrap(true);

  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  sh.getRange(2, 1, 1, width).merge().setFontSize(9).setFontStyle('italic').setFontColor('#8a6d1f')
    .setBackground('#fff7e0').setWrap(true);
  sh.getRange(3, 1).setFontSize(9).setFontColor('#6b7280');

  sh.getRange(headerRow, 1, 1, width).setFontWeight('bold').setBackground('#1f2a44').setFontColor('#ffffff')
    .setVerticalAlignment('middle');
  sh.getRange(1, 1, rows, 1).setFontWeight('bold').setFontColor('#1f2a44');

  // One call for all of them: a tab with twenty sections was twenty round trips
  // to say the same thing twenty times.
  if (groupRows.length) {
    sh.getRangeList(groupRows.map(function (r) {
      return sh.getRange(r, 1, 1, width).getA1Notation();
    })).setBackground('#eef1f6').setFontWeight('bold').setFontColor('#3b4a63').setFontSize(9);
  }

  sh.setColumnWidth(1, 210);
  if (width > 1) sh.setColumnWidths(2, width - 1, 340);
  sh.setFrozenRows(headerRow);
  sh.setFrozenColumns(1);
  if (rows > headerRow) {
    sh.getRange(headerRow, 1, rows - headerRow + 1, width)
      .setBorder(true, true, true, true, true, true, '#d7dce5', SpreadsheetApp.BorderStyle.SOLID);
  }

  // A warning rather than a lock: people need to copy out of this tab, and a
  // dismissible prompt says "your typing will be replaced" better than a
  // permission error does. Re-applied only when it is missing -- removing and
  // recreating a protection on every rebuild is two slow calls to end up where
  // it started.
  var mine = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .filter(function (p) { return p.getDescription() === VIEW_NOTE; });
  if (!mine.length) {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    sh.protect().setDescription(VIEW_NOTE).setWarningOnly(true);
  }
}

function folderByName_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
