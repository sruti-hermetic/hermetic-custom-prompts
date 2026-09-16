/*
 * What the readable view has to keep doing. Run with: sh tests/run.sh
 *
 * These drive the real apps-script.gs against a stand-in for Google's services
 * (tests/apps-script-mock.js), so a change that breaks one of these breaks the
 * spreadsheet ten people are working in. Every case here is one that has either
 * gone wrong or would have: locations with no account row, two customers with
 * the same name, a rebuild that runs out of time, a trigger that piles up, a
 * save landing in the middle of a rebuild.
 */

var pass = 0, fail = 0;
function ok(name, cond, extra){
  if (cond){ pass++; print('  PASS  ' + name); }
  else { fail++; print('  FAIL  ' + name + (extra ? '  -- ' + extra : '')); }
}
function section(t){ print('\n' + t); }

function kv(){ return SS.getSheetByName('kv'); }
function reset(){
  SS.sheets = []; SS.active = null; SS.toasts = [];
  for (var k in PROPS) delete PROPS[k];
  TRIGGERS = []; LOCKS = {};
  for (var c in CALLS) delete CALLS[c];
  var sh = SS.insertSheet('kv'); sh.appendRow(['key','value']);
}
function put(key, obj){ setKey_({ key: key, value: JSON.stringify(obj) }); }
function tabNames(){ return SS.getSheets().map(function(s){ return s.getName(); }); }
function tab(n){ return SS.getSheetByName(n); }
function col(sh, c){
  var out = []; for (var r=1; r<=sh.getLastRow(); r++) out.push(String(sh.cell(r,c)));
  return out;
}
function rowLabels(sh){ return col(sh, 1); }

function venue(accountId, id, name, data, extra){
  var o = { rev:1, updatedAt: Date.now(), createdAt: Date.now(), fieldTs:{},
            lastWriter:'c_test', id:id, accountId:accountId, name:name,
            deleted:false, data: data || {}, docsRemoved:[], lastEditedAt: Date.now(),
            lastWriterName:'Sruti' };
  for (var k in (extra||{})) o[k] = extra[k];
  return o;
}
function account(id, name){
  return { rev:1, updatedAt: Date.now(), createdAt: Date.now(), fieldTs:{},
           lastWriter:'c_test', id:id, name:name, drive_link:'', deleted:false };
}

/* ---------------------------------------------------------------- */
section('1. The spreadsheet as it stands: venues whose account row is missing');
reset();
put('hermetic:presence:c_mu2v8n4p_pjaj9p', {clientId:'c_mu2v8n4p_pjaj9p',name:'Sruti',accountId:'a_mu4g1olr_q61b8',at:1,updatedAt:1,rev:0});
put('hermetic:settings', {rev:17,updatedAt:2,rootFolderName:'ONBOARDING - Custom Prompt Uploads'});
put('hermetic:venue:a_mu4g1olr_q61b8:i_mu4g8f64_2zyb5',
    venue('a_mu4g1olr_q61b8','i_mu4g8f64_2zyb5','Chicago',
      {drive_link:'https://drive.google.com/x', overview:'Modern American restaurant in the West Loop.',
       pricing_text:'', fees_text:'Service charge: 22%.',
       docs:[{name:'test-menu.jpeg',category:'Menu',fileId:'1iD'}]},
      {lastCopiedAt: 3}));
put('hermetic:venue:a_mu4g1olr_q61b8:i_mu4g92t2_8e3eg',
    venue('a_mu4g1olr_q61b8','i_mu4g92t2_8e3eg','San Antonio', {drive_link:'', docs:[]}));

ok('a save marked the account dirty', PROPS['hermetic.dirty.a_mu4g1olr_q61b8'] !== undefined);
ok('a save installed the standing trigger', TRIGGERS.length === 1 &&
   TRIGGERS[0].getHandlerFunction() === 'drainViewQueue', 'triggers=' + TRIGGERS.length);
ok('presence did not mark anything dirty', Object.keys(readDirty_(PropertiesService.getDocumentProperties()))
   .indexOf('hermetic:presence') === -1);

drainViewQueue();
var unfiled = tab('Unfiled locations');
ok('a tab was built for the orphaned locations', !!unfiled, 'tabs=' + tabNames().join('|'));
if (unfiled){
  ok('both locations are columns', String(unfiled.cell(5,2)) === 'Chicago' &&
     String(unfiled.cell(5,3)) === 'San Antonio',
     unfiled.cell(5,2) + ' / ' + unfiled.cell(5,3));
  ok('the answers are on the rows', rowLabels(unfiled).indexOf('Overview') !== -1,
     rowLabels(unfiled).join(' | '));
  ok('an answer reached its cell',
     col(unfiled, 2).join('\n').indexOf('Modern American restaurant') !== -1);
  ok('files are listed readably',
     col(unfiled, 2).join('\n').indexOf('- Menu: test-menu.jpeg') !== -1);
  ok('the tab says why these are unfiled',
     String(unfiled.cell(2,1)).indexOf('no account record') !== -1);
}
ok('kv was hidden behind the readable tab', kv().isSheetHidden());
ok('the dirty list was cleared', Object.keys(readDirty_(PropertiesService.getDocumentProperties())).length === 0,
   JSON.stringify(readDirty_(PropertiesService.getDocumentProperties())));

/* ---------------------------------------------------------------- */
section("2. The account row arrives: the tab takes the customer's name");
put('hermetic:account:a_mu4g1olr_q61b8', account('a_mu4g1olr_q61b8', "Calloway's West Loop"));
drainViewQueue();
ok('the tab was renamed, not duplicated', !!tab("Calloway's West Loop") && !tab('Unfiled locations'),
   tabNames().join('|'));
ok('the locations came with it', tab("Calloway's West Loop") &&
   String(tab("Calloway's West Loop").cell(5,2)) === 'Chicago');
ok('no orphan note on a real account',
   String(tab("Calloway's West Loop").cell(2,1)).indexOf('no account record') === -1);

/* ---------------------------------------------------------------- */
section('3. Two customers with the same name get a tab each');
reset();
put('hermetic:account:a_one', account('a_one', 'The Vineyard'));
put('hermetic:venue:a_one:v1', venue('a_one','v1','Napa', {overview:'FIRST ACCOUNT'}));
put('hermetic:account:a_two', account('a_two', 'The Vineyard'));
put('hermetic:venue:a_two:v2', venue('a_two','v2','Sonoma', {overview:'SECOND ACCOUNT'}));
drainViewQueue();
var vines = tabNames().filter(function(n){ return n.indexOf('The Vineyard') === 0; });
ok('two distinct tabs', vines.length === 2, tabNames().join('|'));
var texts = vines.map(function(n){ return col(tab(n), 2).join('\n'); }).join('\n');
ok("neither account's answers were overwritten",
   texts.indexOf('FIRST ACCOUNT') !== -1 && texts.indexOf('SECOND ACCOUNT') !== -1);

/* ---------------------------------------------------------------- */
section('4. One save rebuilds one tab, not the whole spreadsheet');
for (var c in CALLS) delete CALLS[c];
put('hermetic:venue:a_one:v1', venue('a_one','v1','Napa', {overview:'EDITED'}));
drainViewQueue();
ok('only one tab was cleared', CALLS['clear'] === 1, 'clear=' + CALLS['clear']);
ok('the edit landed', col(tab(readRegistry_()['a_one']), 2).join('\n').indexOf('EDITED') !== -1);
ok('the other tab was left alone',
   col(tab(readRegistry_()['a_two']), 2).join('\n').indexOf('SECOND ACCOUNT') !== -1);
ok('no tab was reordered for a content-only change', CALLS['insertSheet'] === undefined);

/* ---------------------------------------------------------------- */
section('5. An idle minute costs almost nothing');
for (var c2 in CALLS) delete CALLS[c2];
drainViewQueue();
ok('an idle tick reads no sheet data', CALLS['setValues'] === undefined && CALLS['clear'] === undefined);
ok('an idle tick still records that it ran', !!JSON.parse(PROPS['hermetic.viewStatus']).at);

/* ---------------------------------------------------------------- */
section('6. Deleting an account takes its tab with it');
var doomed = readRegistry_()['a_two'];
var dead = account('a_two','The Vineyard'); dead.deleted = true; dead.rev = 2;
put('hermetic:account:a_two', dead);
put('hermetic:venue:a_two:v2', (function(){ var v = venue('a_two','v2','Sonoma',{}); v.deleted = true; return v; })());
drainViewQueue();
ok('the tab is gone', !tab(doomed), tabNames().join('|'));
ok('it left the registry', readRegistry_()['a_two'] === undefined);
ok('the surviving account kept its tab', !!tab(readRegistry_()['a_one']));

/* ---------------------------------------------------------------- */
section('7. A location removed from an account does not break the merged header');
reset();
put('hermetic:account:a_x', account('a_x','Wide'));
put('hermetic:venue:a_x:v1', venue('a_x','v1','One',{overview:'a'}));
put('hermetic:venue:a_x:v2', venue('a_x','v2','Two',{overview:'b'}));
put('hermetic:venue:a_x:v3', venue('a_x','v3','Three',{overview:'c'}));
drainViewQueue();
ok('three columns first time', String(tab('Wide').cell(5,4)) === 'Three');
var gone = venue('a_x','v3','Three',{}); gone.deleted = true;
put('hermetic:venue:a_x:v3', gone);
var threw = null;
try { drainViewQueue(); } catch (e) { threw = e.message; }
var st = JSON.parse(PROPS['hermetic.viewStatus']);
ok('the narrower rebuild did not throw', !threw && !st.error, threw || st.error);
ok('two columns now', String(tab('Wide').cell(5,3)) === 'Two' && String(tab('Wide').cell(5,4)) === '');

/* ---------------------------------------------------------------- */
section('8. The trigger cannot pile up, and comes back if it goes');
reset();
put('hermetic:account:a_t', account('a_t','Trigger'));
ok('one trigger after the first save', TRIGGERS.length === 1, 'n=' + TRIGGERS.length);
for (var i=0;i<40;i++){
  PROPS['hermetic.viewTriggerCheckedAt'] = '0';        // force the re-check every time
  put('hermetic:account:a_t', account('a_t','Trigger ' + i));
}
ok('still exactly one trigger after 40 saves', TRIGGERS.length === 1, 'n=' + TRIGGERS.length);
TRIGGERS = [];                                          // somebody deleted it
PROPS['hermetic.viewTriggerCheckedAt'] = '0';
put('hermetic:account:a_t', account('a_t','Trigger back'));
ok('a save put it back', TRIGGERS.length === 1 &&
   TRIGGERS[0].getHandlerFunction() === 'drainViewQueue');

section('9. Leftovers from the previous scheme are cleared out');
reset();
TRIGGERS.push({ getHandlerFunction: function(){ return 'rebuildViewsQueued'; } });
TRIGGERS.push({ getHandlerFunction: function(){ return 'rebuildViewsQueued'; } });
TRIGGERS.push({ getHandlerFunction: function(){ return 'rebuildViews'; } });
put('hermetic:account:a_o', account('a_o','Old'));
ok('the one-off triggers were reclaimed', TRIGGERS.length === 1 &&
   TRIGGERS[0].getHandlerFunction() === 'drainViewQueue',
   TRIGGERS.map(function(t){ return t.getHandlerFunction(); }).join(','));

section('10. A hand edit in kv reaches the readable tab');
reset();
put('hermetic:account:a_h', account('a_h','Hand'));
put('hermetic:venue:a_h:v1', venue('a_h','v1','Desk',{overview:'typed in the tool'}));
drainViewQueue();
var row = findRow_(kv(), 'hermetic:venue:a_h:v1');
var rec = JSON.parse(readRowValue_(kv(), row));
var before = JSON.stringify(rec);
rec.data.overview = 'typed into the sheet';
writeRow_(kv(), row, 'hermetic:venue:a_h:v1', chunksOf_(JSON.stringify(rec)));
normalizeEditedRow_(kv(), row, { oldValue: before, range: {
  getNumRows: function(){ return 1; }, getNumColumns: function(){ return 1; }, getColumn: function(){ return 2; } } });
ok('the hand edit marked the account dirty', PROPS['hermetic.dirty.a_h'] !== undefined);
drainViewQueue();
ok('and it shows on the tab',
   col(tab('Hand'), 2).join('\n').indexOf('typed into the sheet') !== -1);

section('11. Published section titles are used, and their order kept');
reset();
put('hermetic:schema', { parts: [
  { title: 'Location facts', sections: [{key:'overview',title:'Overview'},{key:'hours_text',title:'Hours and closures'}] },
  { title: 'Assistant behavior', sections: [{key:'persona_text',title:'Persona'}] } ], updatedAt: 1 });
put('hermetic:account:a_s', account('a_s','Schema'));
put('hermetic:venue:a_s:v1', venue('a_s','v1','Main',{overview:'o',hours_text:'h',persona_text:'p'}));
drainViewQueue();
var labels = rowLabels(tab('Schema'));
ok('the published titles are the row labels',
   labels.indexOf('Hours and closures') !== -1 && labels.indexOf('Persona') !== -1, labels.join(' | '));
ok('the published order is kept',
   labels.indexOf('Overview') < labels.indexOf('Hours and closures') &&
   labels.indexOf('LOCATION FACTS') < labels.indexOf('ASSISTANT BEHAVIOR'), labels.join(' | '));

section('12. A schema change rebuilds every account');
put('hermetic:account:a_s2', account('a_s2','Second'));
drainViewQueue();
for (var c3 in CALLS) delete CALLS[c3];
put('hermetic:schema', { parts: [
  { title: 'Location facts', sections: [{key:'overview',title:'Overview RENAMED'}] } ], updatedAt: 2 });
drainViewQueue();
ok('both tabs were rebuilt', CALLS['clear'] === 2, 'clear=' + CALLS['clear']);
ok('the new title is on the tab', rowLabels(tab('Schema')).indexOf('Overview RENAMED') !== -1);

section('13. ?action=diag reports what is actually wrong');
reset();
put('hermetic:venue:a_missing:v1', venue('a_missing','v1','Orphan',{}));
var diag = JSON.parse(doGet({ parameter: { action: 'diag' } })._body);
ok('diag counts the orphaned locations', diag.rows.orphanVenues === 1, JSON.stringify(diag.rows));
ok('diag reports no schema row yet', diag.rows.schema === false);
ok('diag reports the trigger', diag.view.trigger === true);
ok('diag lists what is waiting', diag.view.pending.length === 1, JSON.stringify(diag.view.pending));
ok('serverVersion moved', diag.serverVersion === 8);

section('14. A save that lands mid-rebuild is not swallowed');
reset();
put('hermetic:account:a_r', account('a_r','Race'));
put('hermetic:venue:a_r:v1', venue('a_r','v1','One',{overview:'first'}));
// The rebuild reads the mark, then someone saves again before it clears it.
var props = PropertiesService.getDocumentProperties();
var seen = props.getProperty('hermetic.dirty.a_r');
props.setProperty('hermetic.dirty.a_r', String(Number(seen) + 1000));   // a newer save
clearDirty_(props, 'a_r', seen);
ok('the newer mark survives the older rebuild clearing it',
   props.getProperty('hermetic.dirty.a_r') !== null);
clearDirty_(props, 'a_r', props.getProperty('hermetic.dirty.a_r'));
ok('clearing with the current stamp does remove it',
   props.getProperty('hermetic.dirty.a_r') === null);

section('15. Two rebuilds cannot run over each other');
reset();
put('hermetic:account:a_l', account('a_l','Lock'));
LOCKS['document'] = true;                        // a rebuild is already running
for (var c in CALLS) delete CALLS[c];
drainViewQueue();
ok('the second one stood down', CALLS['clear'] === undefined && CALLS['prop.all'] === undefined);
ok('and left the work on the list', PROPS['hermetic.dirty.a_l'] !== undefined);
LOCKS['document'] = false;
drainViewQueue();
ok('the next tick picks it up', !!tab('Lock'));

section('16. More accounts than fit in one run still all get built');
reset();
for (var i=0;i<8;i++){
  put('hermetic:account:a_' + i, account('a_' + i, 'Customer ' + i));
  put('hermetic:venue:a_' + i + ':v', venue('a_' + i, 'v', 'Main', {overview:'answer ' + i}));
}
var realBudget = VIEW_BUDGET_MS;
VIEW_BUDGET_MS = -1;                             // every account is over budget at once
drainViewQueue();
var afterFirst = Object.keys(readRegistry_()).length;
ok('the run stopped early rather than dying', afterFirst >= 1 && afterFirst < 8, 'built=' + afterFirst);
ok('it said how much was left', JSON.parse(PROPS['hermetic.viewStatus']).left > 0);
var guard = 0;
while (Object.keys(readDirty_(PropertiesService.getDocumentProperties())).length && guard++ < 50) drainViewQueue();
VIEW_BUDGET_MS = realBudget;
ok('later ticks finished the rest', Object.keys(readRegistry_()).length === 8,
   'built=' + Object.keys(readRegistry_()).length + ' ticks=' + guard);
ok('every customer has a tab', tabNames().filter(function(n){ return n.indexOf('Customer ') === 0; }).length === 8);
ok('nothing was built twice', tabNames().length === new Set(tabNames()).size);

section('17. A refresh that fails never costs somebody their answer');
reset();
var realProps = PropertiesService.getDocumentProperties;
PropertiesService.getDocumentProperties = function(){ throw new Error('properties unavailable'); };
var res = null, threw = null;
try { res = JSON.parse(setKey_({ key:'hermetic:venue:a_q:v1', value: JSON.stringify(venue('a_q','v1','Q',{overview:'kept'})) })._body); }
catch (e) { threw = e.message; }
PropertiesService.getDocumentProperties = realProps;
ok('the save still succeeded', !threw && res && res.ok === true, threw || JSON.stringify(res));
ok('and the answer is in kv', readRowValue_(kv(), findRow_(kv(),'hermetic:venue:a_q:v1')).indexOf('kept') !== -1);

section('18. Ten people saving: what one save costs');
reset();
for (var a=0;a<10;a++){
  put('hermetic:account:b_' + a, account('b_' + a, 'Account ' + a));
  for (var v=0;v<5;v++) put('hermetic:venue:b_' + a + ':v' + v, venue('b_'+a,'v'+v,'Venue '+v,{overview:'x'}));
}
drainViewQueue();
ok('all ten tabs exist', Object.keys(readRegistry_()).length === 10);
for (var c2 in CALLS) delete CALLS[c2];
put('hermetic:venue:b_3:v2', venue('b_3','v2','Venue 2',{overview:'one word changed'}));
var saveCalls = (CALLS['prop.get']||0) + (CALLS['prop.set']||0) + (CALLS['trigger.list']||0);
ok('a save adds only a couple of cheap calls', saveCalls <= 4, 'calls=' + saveCalls);
ok('a save never touches a trigger', CALLS['trigger.create'] === undefined && CALLS['trigger.list'] === undefined);
for (var c3 in CALLS) delete CALLS[c3];
drainViewQueue();
ok('the rebuild touched one tab of ten', CALLS['clear'] === 1, 'clear=' + CALLS['clear']);
ok('it did not reshuffle the spreadsheet', CALLS['deleteSheet'] === undefined && CALLS['insertSheet'] === undefined);
var whole = (function(sh){ var t=[]; for (var r=1;r<=sh.getLastRow();r++)
  for (var c=1;c<=6;c++) t.push(String(sh.cell(r,c))); return t.join('\n'); })(tab(readRegistry_()['b_3']));
ok('the change is visible', whole.indexOf('one word changed') !== -1);
ok('and only on that account', (function(sh){ var t=[]; for (var r=1;r<=sh.getLastRow();r++)
  for (var c=1;c<=6;c++) t.push(String(sh.cell(r,c))); return t.join('\n'); })(tab(readRegistry_()['b_4']))
  .indexOf('one word changed') === -1);

section('19. Upgrading from the previous version in place');
reset();
// What the old build left behind: tabs already on the sheet, a registry of bare
// names, a spent one-off trigger, and no dirty list at all.
SS.insertSheet('Account 0').appendRow(['stale contents']);
SS.insertSheet('Account 1').appendRow(['stale contents']);
PROPS['hermetic.viewTabs'] = JSON.stringify(['Account 0','Account 1']);
PROPS['hermetic.viewRebuildQueuedAt'] = String(Date.now());
TRIGGERS.push({ getHandlerFunction: function(){ return 'rebuildViewsQueued'; } });
put('hermetic:account:c_0', account('c_0','Account 0'));
put('hermetic:venue:c_0:v', venue('c_0','v','Main',{overview:'live answer'}));
put('hermetic:account:c_1', account('c_1','Account 1'));
put('hermetic:venue:c_1:v', venue('c_1','v','Main',{overview:'other answer'}));
drainViewQueue();
ok('the old tabs were adopted, not duplicated', tabNames().filter(function(n){
  return n.indexOf('Account 0') === 0; }).length === 1, tabNames().join('|'));
ok('their stale contents were replaced',
   col(tab('Account 0'), 2).join('\n').indexOf('live answer') !== -1);
ok('the registry is keyed by account now',
   readRegistry_()['c_0'] === 'Account 0' && readRegistry_()['c_1'] === 'Account 1',
   JSON.stringify(readRegistry_()));

section('20. Renaming a customer renames the tab');
put('hermetic:account:c_0', (function(){ var a = account('c_0','Account 0 Renamed'); a.rev = 2; return a; })());
drainViewQueue();
ok('the tab followed the name', !!tab('Account 0 Renamed') && !tab('Account 0'), tabNames().join('|'));
ok('the answers came with it', col(tab('Account 0 Renamed'), 2).join('\n').indexOf('live answer') !== -1);
ok('tabs are in alphabetical order',
   tabNames().filter(function(n){ return n !== 'kv'; }).join('|') === 'Account 0 Renamed|Account 1',
   tabNames().join('|'));

section('21. A row half written when the rebuild reads it');
reset();
put('hermetic:account:d_0', account('d_0','Fine'));
put('hermetic:venue:d_0:v', venue('d_0','v','Main',{overview:'good'}));
put('hermetic:account:d_1', account('d_1','Broken'));
// a value caught mid-write: valid row, unparseable JSON
writeRow_(kv(), findRow_(kv(),'hermetic:account:d_1'), 'hermetic:account:d_1', ['{"rev":2,"na']);
var threw2 = null;
try { drainViewQueue(); } catch (e) { threw2 = e.message; }
ok('the rebuild did not die on it', !threw2, threw2);
ok('the healthy account still got its tab', !!tab('Fine'));


print('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) throw new Error(fail + ' failing');
