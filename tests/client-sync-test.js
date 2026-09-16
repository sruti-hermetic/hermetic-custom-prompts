/*
 * What the console must keep sending to the Sheet. Run with: sh tests/run.sh
 *
 * These run the real queueLocalOnly out of index.html -- run.sh cuts it and the
 * handful of helpers it uses straight out of the file -- against a stub
 * localStorage. The case that matters is the one that actually happened: an
 * account created before the sync URL was pasted stayed in one browser, so the
 * spreadsheet had that account's locations and no account to file them under.
 */
/* Stubs for the browser bits queueLocalOnly leans on. */
var STORE = {};
var localStorage = {
  get length(){ return Object.keys(STORE).length; },
  key: function(i){ return Object.keys(STORE)[i]; },
  getItem: function(k){ return STORE[k] === undefined ? null : STORE[k]; },
  setItem: function(k,v){ STORE[k] = String(v); },
  removeItem: function(k){ delete STORE[k]; }
};
var console = { warn: function(m){ print('   [warn] ' + m); }, log: function(){}, error: function(){} };
var writeQueue = new Map();
var syncUrl = 'https://exec';
var persisted = 0, wroteNow = 0;
function persistDirty(){ persisted++; }
function writeNow(){ wroteNow++; }
load('UNDER_TEST');   // spliced in by run.sh

var pass = 0, fail = 0;
function ok(n,c,e){ if(c){pass++;print('  PASS  '+n);} else {fail++;print('  FAIL  '+n+(e?'  -- '+e:''));} }
function section(t){ print('\n'+t); }
function reset(){ STORE = {}; writeQueue = new Map(); persisted = 0; wroteNow = 0; }
function rec(o){ return JSON.stringify(Object.assign({rev:1,updatedAt:1,fieldTs:{}}, o)); }

section('The account that never left the laptop');
reset();
// what Sruti's browser looks like: an account made before the URL was pasted,
// and venues saved after it was
STORE['hermetic:account:a_riverside'] = rec({id:'a_riverside', name:'Riverside Hospitality Group', deleted:false});
STORE['hermetic:venue:a_riverside:i_chi'] = rec({id:'i_chi', accountId:'a_riverside', name:'Chicago'});
STORE['hermetic:venue:a_riverside:i_sat'] = rec({id:'i_sat', accountId:'a_riverside', name:'San Antonio'});
var fromSheet = [
  { key:'hermetic:venue:a_riverside:i_chi', value: STORE['hermetic:venue:a_riverside:i_chi'] },
  { key:'hermetic:venue:a_riverside:i_sat', value: STORE['hermetic:venue:a_riverside:i_sat'] },
  { key:'hermetic:settings', value: rec({rootFolderName:'ONBOARDING'}) },
  { key:'hermetic:presence:c_other', value: rec({clientId:'c_other'}) }
];
var n = queueLocalOnly(fromSheet);
ok('the missing account was queued', n === 1 && writeQueue.has('hermetic:account:a_riverside'), 'n='+n);
ok('the venues already up were left alone', writeQueue.size === 1,
   [...writeQueue.keys()].join('|'));
ok('it was sent at once, not on the next window', wroteNow === 1 && persisted === 1);
ok('it was queued as it stands, not re-saved',
   writeQueue.get('hermetic:account:a_riverside').rev === 1);

section('A second load has nothing left to send');
var nowUp = fromSheet.concat([{ key:'hermetic:account:a_riverside', value: STORE['hermetic:account:a_riverside'] }]);
writeQueue = new Map(); wroteNow = 0;
ok('nothing is queued twice', queueLocalOnly(nowUp) === 0 && wroteNow === 0);

section('What must never be pushed');
reset();
STORE['hermetic:presence:c_me'] = rec({clientId:'c_me'});          // heartbeat owns this
STORE['hermetic:schema'] = JSON.stringify({parts:[]});             // publishSchema owns this
STORE['hermetic:accounts'] = JSON.stringify([{id:'a_old'}]);       // legacy list, folded not uploaded
STORE['hermetic:venues:a_old'] = JSON.stringify([{id:'v'}]);       // legacy list
STORE['hermetic:syncUrl'] = 'https://exec';                        // never leaves the browser
STORE['hermetic:displayName'] = 'Sruti';
STORE['hermetic:dirtyKeys'] = '[]';
STORE['hermetic:clientId'] = 'c_me';
ok('none of it is queued', queueLocalOnly([]) === 0, [...writeQueue.keys()].join('|'));

section('Settings and tombstones');
reset();
STORE['hermetic:settings'] = rec({rootFolderName:'ONBOARDING'});
STORE['hermetic:account:a_gone'] = rec({id:'a_gone', deleted:true});
ok('local-only settings go up', queueLocalOnly([]) === 2);
ok('a tombstone the Sheet lacks goes up too', writeQueue.has('hermetic:account:a_gone'));
reset();
STORE['hermetic:account:a_gone'] = rec({id:'a_gone', deleted:true});
ok('a tombstone the Sheet already has is left alone',
   queueLocalOnly([{key:'hermetic:account:a_gone', value:STORE['hermetic:account:a_gone']}]) === 0);

section('Guards');
reset();
STORE['hermetic:account:a_bad'] = 'not json';
STORE['hermetic:account:a_arr'] = '[1,2,3]';
STORE['hermetic:account:a_ok'] = rec({id:'a_ok'});
ok('unreadable rows are skipped, the good one still goes', queueLocalOnly([]) === 1 &&
   writeQueue.has('hermetic:account:a_ok'), [...writeQueue.keys()].join('|'));
reset();
syncUrl = '';
STORE['hermetic:account:a_x'] = rec({id:'a_x'});
ok('nothing is queued with no Sheet configured', queueLocalOnly([]) === 0 && wroteNow === 0);
syncUrl = 'https://exec';

print('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) throw new Error(fail + ' failing');
