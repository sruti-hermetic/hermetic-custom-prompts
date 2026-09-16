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

section('The sections the console collects');
var locKeys = LOCATION_SCHEMA.map(function(b){ return b.key; });
ok('Policies is a box', locKeys.indexOf('policies_text') !== -1, locKeys.join(','));
ok('FAQs is a box', locKeys.indexOf('faqs_text') !== -1, locKeys.join(','));
ok('sub-locations is gone', locKeys.indexOf('sub_venues_text') === -1, locKeys.join(','));
ok('Policies and FAQs sit after Booking rules, as the template orders them',
   locKeys.indexOf('booking_text') < locKeys.indexOf('policies_text') &&
   locKeys.indexOf('policies_text') < locKeys.indexOf('faqs_text'));
ok('every box has guidance to fill it by',
   LOCATION_SCHEMA.every(function(b){ return b.hint && b.hint.length; }));
var blank = newVenueData();
ok('a new location starts with both boxes',
   blank.policies_text === '' && blank.faqs_text === '', Object.keys(blank).join(','));
ok('and with no sub-locations box', blank.sub_venues_text === undefined);

section('What Claude is handed');
var filled = newVenueData();
filled.policies_text = 'Outside food / catering: not permitted, except a celebration cake.';
filled.faqs_text = 'Can we bring a cake? Outside desserts such as wedding cakes are permitted.';
var out = serializePart(LOCATION_SCHEMA, filled);
ok('Policies is in the prompt', out.indexOf('Policies:\nOutside food / catering') !== -1);
ok('FAQs is in the prompt', out.indexOf('FAQs:\nCan we bring a cake?') !== -1);
ok('nothing offers sub-venues any more', out.toLowerCase().indexOf('sub-location') === -1 &&
   out.toLowerCase().indexOf('sub-venue') === -1);
ok('an empty box still reports itself, so Claude knows it was asked',
   serializePart(LOCATION_SCHEMA, newVenueData()).indexOf('Policies:\nNone.') !== -1);
ok('the skill template no longer asks for Sub-Venues',
   SKILL_TEXT.indexOf('## Sub-Venues') === -1);
ok('the skill template still asks for Policies and FAQs',
   SKILL_TEXT.indexOf('## Policies') !== -1 && SKILL_TEXT.indexOf('## FAQs') !== -1);
ok('Claude is no longer told policies come only from the files',
   FILE_SOURCED_LABEL.toLowerCase().indexOf('polic') === -1 &&
   FILE_SOURCED_LABEL.toLowerCase().indexOf('faq') === -1, FILE_SOURCED_LABEL);

section('Rows typed before these were boxes');
var legacy = { policies: [{topic:'Outside food', policy:'Not permitted.'},
                          {topic:'Open flame', policy:'Candles must be enclosed.'}],
               faqs: [{question:'Can we see the space?', answer:'Yes, during business hours.'}],
               sub_venues: [{name:'The Grill', description:'Casual counter.'}] };
foldLegacyBlocks(legacy);
ok('old policy rows became the box text',
   legacy.policies_text.indexOf('Outside food') !== -1 &&
   legacy.policies_text.indexOf('Candles must be enclosed') !== -1, legacy.policies_text);
ok('old FAQ rows became the box text',
   legacy.faqs_text.indexOf('Can we see the space?') !== -1, legacy.faqs_text);
ok('the rows themselves were dropped, so nothing prints twice',
   legacy.policies === undefined && legacy.faqs === undefined && legacy.sub_venues === undefined);
var typed = { policies_text: 'What the teammate typed.', policies: [{topic:'Old', policy:'Older.'}] };
foldLegacyBlocks(typed);
ok('a box somebody has already filled is never overwritten',
   typed.policies_text === 'What the teammate typed.');
ok('sub-locations is on the retired list, so it is cleaned out of saved records',
   RETIRED_FIELDS.indexOf('sub_venues_text') !== -1);

section('Where a guest-count-to-space rule belongs');
function box(schema, key){ return schema.filter(function(b){ return b.key === key; })[0]; }
var avail = box(LOCATION_SCHEMA, 'availability_text').hint.join(' ').toLowerCase();
var rules = box(CAMPAIGN_SCHEMA, 'rules_text').hint.join(' ').toLowerCase();
ok('Availability no longer asks which room suits which guest count',
   avail.indexOf('which rooms may be offered') === -1);
ok('Availability asks about the venue instead',
   avail.indexOf('one booked space blocks the whole building') !== -1, avail);
ok('and points the guest-count question at Rules', avail.indexOf('rules') !== -1, avail);
ok('Rules asks for the guest-count bands',
   rules.indexOf('which space to offer at which guest count') !== -1, rules);
ok('Rules says an offer names one space',
   rules.indexOf('no combining two into one offer') !== -1, rules);

section('And what Claude is told about it');
ok('the template says Availability is not about which space to offer',
   SKILL_TEXT.indexOf('not which space to offer') !== -1);
ok('the template sends routing to a Campaign Rule',
   SKILL_TEXT.indexOf('Guest-count-to-space routing is a Campaign Rule') !== -1);
ok('the placement test agrees',
   SKILL_TEXT.indexOf('Which space to offer at a guest count is a Rule, never an Availability Policy line') !== -1);
ok('a capacity is still a Location fact',
   SKILL_TEXT.indexOf("A room's capacity is a fact (Location \u2192 Rooms)") !== -1);
ok('Mia is told she has no combined spaces',
   SKILL_TEXT.indexOf('no concept of a combined or shared space') !== -1);
ok('a paired offer goes to Open Questions rather than being invented',
   SKILL_TEXT.indexOf('put the pairing in Open Questions') !== -1);
ok('a rule that cannot be applied goes to Open Questions',
   SKILL_TEXT.indexOf('cannot be applied without deciding what it means') !== -1);
ok('with the whole-evening line as the worked example',
   SKILL_TEXT.indexOf('Saturday events are whole-evening only') !== -1);
ok('the self-check catches a two-space offer',
   SKILL_TEXT.indexOf('Any offer naming two spaces at once') !== -1);
ok('the self-check catches an unapplicable rule',
   SKILL_TEXT.indexOf('without deciding what it meant') !== -1);

print('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) throw new Error(fail + ' failing');
