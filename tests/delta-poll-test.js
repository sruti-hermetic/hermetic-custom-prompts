/*
 * What a teammate's open console must keep receiving. Run with: sh tests/run.sh
 *
 * Drives the real apps-script.gs. The case that happened: a teammate created
 * eight venues, and everyone else's console showed them one at a time over
 * hours. Each write sat in the console's 30s window, landed after other
 * people had already polled past the moment it was typed, and every delta
 * poll after that skipped it until they reloaded.
 */
var T = 1000000;
Date.now = function(){ return T; };
var s = 1000, PAD = 5000;   // index.html's CLOCK_PAD_MS

/* CacheService with a real TTL on the fake clock, so the cache race is testable. */
var CACHE = {};
var CacheService = { getDocumentCache: function(){ return {
  get: function(k){ var e = CACHE[k]; return e && e.until > T ? e.v : null; },
  put: function(k, v, ttl){ CACHE[k] = { v: v, until: T + ttl * 1000 }; },
  remove: function(k){ delete CACHE[k]; }
}; } };

var pass = 0, fail = 0;
function ok(n,c,e){ if(c){pass++;print('  PASS  '+n);} else {fail++;print('  FAIL  '+n+(e?'  -- '+e:''));} }
function section(t){ print('\n'+t); }
function kv(){ return SS.getSheetByName('kv'); }
function reset(){
  SS.sheets = []; CACHE = {}; for (var k in PROPS) delete PROPS[k]; TRIGGERS = [];
  kv() || SS.insertSheet('kv').appendRow(['key','value']);
}
function set(key, value, ifRev){
  var body = { action: 'set', key: key, value: typeof value === 'string' ? value : JSON.stringify(value) };
  if (ifRev !== undefined) body.ifRev = ifRev;
  return JSON.parse(doPost({ postData: { contents: JSON.stringify(body) } })._body);
}
function getAll(since){
  return JSON.parse(doGet({ parameter: { action: 'getAll', prefix: 'hermetic:', since: since ? String(since) : '' } })._body);
}
/* A console's poll loop, cursor and all, exactly as index.html keeps it. */
function Console(){ this.cursor = 0; this.seen = {}; }
Console.prototype.pull = function(){
  var r = getAll(this.cursor);
  var self = this;
  r.items.forEach(function(it){ self.seen[it.key] = it.value; });
  this.cursor = r.now - PAD;
  return r.items;
};
function venue(id, name, typedAt){
  return { id: id, accountId: 'a1', name: name, rev: 1, updatedAt: typedAt, createdAt: typedAt,
           fieldTs: { name: typedAt }, data: {} };
}

section('1. A venue that sat in the 30s window still reaches an open console');
reset();
T = 100*s; set('hermetic:account:a1', { id:'a1', name:'Acme', rev:1, updatedAt:T, fieldTs:{} });
var b = new Console(); b.pull();
var typedAt = 110*s;
T = 120*s; b.pull();                                   // polls before the write lands
T = 135*s; set('hermetic:venue:a1:i1', venue('i1','Downtown', typedAt));
T = 140*s; b.pull();
ok('the next poll brings it', !!b.seen['hermetic:venue:a1:i1']);

section('2. Eight venues created in a burst all arrive, with no reload');
reset();
T = 100*s; set('hermetic:account:a1', { id:'a1', name:'Acme', rev:1, updatedAt:T, fieldTs:{} });
b = new Console(); b.pull();
var burst = [];
for (var i = 0; i < 8; i++) burst.push(venue('i'+i, 'Venue '+i, (101 + i)*s));
T = 115*s; b.pull(); T = 135*s; b.pull();             // two polls while all eight wait
T = 140*s; burst.forEach(function(v){ T += 1200; set('hermetic:venue:a1:'+v.id, v); });
T = 160*s; b.pull(); T = 180*s; b.pull();
var got = burst.filter(function(v){ return b.seen['hermetic:venue:a1:'+v.id]; }).length;
ok('all eight are on the open console', got === 8, got + ' of 8');

section('3. A laptop whose clock is 10 minutes slow is not invisible');
reset();
b = new Console(); T = 100*s; b.pull();
T = 120*s; set('hermetic:venue:a1:slow', venue('slow','Slow clock', T - 10*60*s));
T = 130*s; b.pull();
ok('its write arrives on the next poll', !!b.seen['hermetic:venue:a1:slow']);

section('4. Clients get back exactly what they sent -- no stamp leaks out');
reset();
var sent = JSON.stringify(venue('i1','Exact', 50*s));
T = 200*s; set('hermetic:venue:a1:i1', sent);
ok('the stored cell carries the stamp',
   String(kv().getRange(2,2,1,1).getValues()[0][0]).indexOf('{"_arrivedAt":200000,') === 0);
var full = getAll(0).items.filter(function(it){ return it.key === 'hermetic:venue:a1:i1'; })[0];
ok('a full read returns it byte for byte', full && full.value === sent, full && full.value);
var delta = getAll(150*s).items.filter(function(it){ return it.key === 'hermetic:venue:a1:i1'; })[0];
ok('a delta read returns it byte for byte', delta && delta.value === sent);
var one = JSON.parse(doGet({ parameter: { action: 'get', key: 'hermetic:venue:a1:i1' } })._body);
ok('get returns it byte for byte', one.value === sent);
var c = set('hermetic:venue:a1:i1', venue('i1','Stale', 60*s), 7);
ok('a conflict hands back the value without the stamp',
   c.conflict === true && c.current === sent, c.current);
var empty = '{}';
T = 210*s; set('hermetic:presence:x', empty);
ok('an empty object round-trips', getAll(0).items.some(function(it){ return it.key === 'hermetic:presence:x' && it.value === '{}'; }));

section('5. Values that are not records are stored as sent');
reset();
T = 300*s; set('hermetic:accounts', '[{"id":"a1"}]');
ok('a legacy list is untouched', kv().getRange(2,2,1,1).getValues()[0][0] === '[{"id":"a1"}]');
ok('and returned on a full read', getAll(0).items[0].value === '[{"id":"a1"}]');

section('6. Rows saved before the upgrade still poll on updatedAt');
reset();
writeRow_(kv(), 2, 'hermetic:venue:a1:old', chunksOf_(JSON.stringify(venue('old','Old', 500*s))));
CACHE = {};
T = 600*s;
ok('newer than the cursor: returned', getAll(400*s).items.length === 1);
ok('older than the cursor: skipped', getAll(550*s).items.length === 0);

section('7. A cached snapshot does not move the cursor past a write it missed');
reset();
b = new Console(); T = 100*s; b.pull();
T = 120*s; var snap = readSnapshot_();                 // read starts and is cached
// a write stamped half a second before that read, not yet flushed when it ran
T = 119*s + 500; set('hermetic:venue:a1:race', venue('race','Race', 119*s));
// the in-flight read finishes after the write and puts its stale snapshot back
T = 120*s + 200;
CacheService.getDocumentCache().put(KV_CACHE_KEY, JSON.stringify(snap), KV_CACHE_TTL_S);
T = 124*s + 900; b.pull();                             // served from the stale cache
ok('the stale snapshot did not carry the write', !b.seen['hermetic:venue:a1:race']);
T = 145*s; b.pull();
ok('the next poll still catches it', !!b.seen['hermetic:venue:a1:race']);

section('8. A hand edit in kv reaches open consoles');
reset();
T = 100*s; set('hermetic:venue:a1:h', venue('h','Hand', 100*s));
b = new Console(); T = 110*s; b.pull();
var row = findRow_(kv(), 'hermetic:venue:a1:h');
var before = readRowValue_(kv(), row);
var rec = JSON.parse(before); rec.name = 'Typed in the sheet';
writeRow_(kv(), row, 'hermetic:venue:a1:h', chunksOf_(JSON.stringify(rec)));
T = 130*s;
normalizeEditedRow_(kv(), row, { oldValue: before, range: {
  getNumRows: function(){ return 1; }, getNumColumns: function(){ return 1; }, getColumn: function(){ return 2; } } });
T = 140*s; b.pull();
var seenRec = JSON.parse(b.seen['hermetic:venue:a1:h']);
ok('the edit arrives on the next poll', seenRec.name === 'Typed in the sheet');
ok('without the stamp', seenRec._arrivedAt === undefined);
ok('and the cell holds exactly one stamp',
   readRowValue_(kv(), row).split('_arrivedAt').length === 2);

section('9. A record longer than one cell keeps its stamp and its tail');
reset();
var big = venue('big','Big', 100*s); big.data.notes = new Array(120001).join('x');
var bigJson = JSON.stringify(big);
T = 200*s; set('hermetic:venue:a1:big', bigJson);
ok('it spans several cells', kv().getRange(2,4,1,1).getValues()[0][0] !== '');
var back = getAll(150*s).items[0];
ok('a delta read returns it whole', back && back.value === bigJson, back && back.value.length + ' vs ' + bigJson.length);

section('10. The readable view still reads stamped rows');
reset();
T = 100*s; set('hermetic:account:a1', { id:'a1', name:'Acme', rev:1, updatedAt:T, fieldTs:{} });
set('hermetic:venue:a1:v', venue('v','Midtown', T));
drainViewQueue();
var tab = SS.getSheetByName('Acme');
ok('the venue is a column on the account tab', tab && tab.getRange(5,2,1,1).getValues()[0][0] === 'Midtown');

print('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) throw new Error(fail + ' failing');
