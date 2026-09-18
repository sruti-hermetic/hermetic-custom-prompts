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
ok('Availability asks what closes the whole venue instead',
   avail.indexOf('what closes the whole venue for a day') !== -1, avail);
ok('Rules asks for the guest-count bands',
   rules.indexOf('which space to offer at which guest count') !== -1, rules);
ok('Rules says an offer names one space',
   rules.indexOf('no combining two into one offer') !== -1, rules);

section('Workflow steps is the whole goal path');
var steps = box(CAMPAIGN_SCHEMA, 'steps_text');
var stepHint = steps.hint.join(' ').toLowerCase();
ok('the box no longer asks only for what differs',
   stepHint.indexOf('only steps that differ') === -1, stepHint);
ok('it asks for every step in order', stepHint.indexOf('every step in order') !== -1, stepHint);
ok('it starts after the name and ends at the handoff',
   stepHint.indexOf('from after the name is confirmed to the handoff') !== -1, stepHint);
ok('it asks for the shape of a step',
   stepHint.indexOf('a name, when it starts as a condition') !== -1 &&
   stepHint.indexOf('what skips it') !== -1, stepHint);
ok('a trigger is a condition, not a step number',
   stepHint.indexOf('not a step number') !== -1, stepHint);
ok('it asks what a step holds back, not only what it asks',
   stepHint.indexOf('holds back as well as asks') !== -1, stepHint);
ok('the guidance is short enough to read at a glance', steps.hint.length <= 3,
   steps.hint.length + ' bullets');
// The worked example is where the branch and the skip now live, since the
// bullets no longer spell them out.
ok('the worked example shows trigger, do and skip',
   (steps.placeholder || '').indexOf('Trigger:') !== -1 &&
   (steps.placeholder || '').indexOf('Do:') !== -1 &&
   (steps.placeholder || '').indexOf('Skip if:') !== -1, steps.placeholder);
ok('and shows a branch rather than a single clean path',
   (steps.placeholder || '').indexOf('If two fit, describe both') !== -1, steps.placeholder);
ok('and stays short enough to read past',
   (steps.placeholder || '').length < 170, (steps.placeholder || '').length + ' chars');

section('The scheduling link is a setting, not a link to type');
var links = box(CAMPAIGN_SCHEMA, 'links_text').hint.join(' ').toLowerCase();
ok('the Links box says so',
   links.indexOf('not the coordinator scheduling link') !== -1, links);
ok('and says where it comes from instead',
   links.indexOf('account settings') !== -1, links);

section('Vocabulary says how far a substitution reaches');
var vocab = box(CAMPAIGN_SCHEMA, 'vocabulary_text');
ok('it is still substitutions only',
   vocab.hint.join(' ').toLowerCase().indexOf('not rules') !== -1);
ok('and now says each one holds always by default',
   vocab.hint.join(' ').toLowerCase().indexOf('holds always, unless you say where it does not') !== -1,
   vocab.hint.join(' '));
ok('the example shows an exception being stated',
   (vocab.placeholder || '').indexOf('except when the guest names a room') !== -1, vocab.placeholder);

section('Required wording is obligation, not preference');
var wording = box(CAMPAIGN_SCHEMA, 'wording_text').hint.join(' ').toLowerCase();
ok('the box leads with the obligation test',
   wording.indexOf('legally or contractually obliged to say') !== -1, wording);
ok('and rules scripts out', wording.indexOf('not a script') !== -1, wording);

section('What a real availability policy needs to say');
ok('room relationships the CRM cannot express, not a substitute for correct nesting',
   avail.indexOf('room relationships the booking system cannot already express') !== -1, avail);
ok('parent/child and combined-space blocking is named',
   avail.indexOf('parent/child rooms, a combined space that blocks its children') !== -1, avail);
ok('turnover buffers and duration caps are asked for',
   avail.indexOf('turnover buffers, duration caps per room') !== -1, avail);
ok('spaces that both fit get a tie-break',
   avail.indexOf('which space to offer when more than one fits') !== -1, avail);
ok('the worked example is the real pattern: a buyout blocking the whole day',
   (box(LOCATION_SCHEMA, 'availability_text').placeholder || '').indexOf('even outside its hours') !== -1);

section('What Drive says happened to the files');
var docA = {name:'Menu.pdf', link:'https://old-link', category:'Menu', fileId:'f_a'};
var docB = {name:'Pricing.pdf', link:'https://link-b', category:'Pricing doc/matrix', fileId:'f_b'};
ok('a file not asked about is left exactly as it was',
   applyFileCheck([docA], {}).docs[0] === docA);
var deleted = applyFileCheck([docA], {f_a:{exists:false}});
ok('a file Drive no longer has is marked gone, not dropped',
   deleted.changed && deleted.docs[0].gone === 'deleted' && deleted.docs.length === 1);
var trashed = applyFileCheck([docA], {f_a:{exists:true, trashed:true}});
ok('a trashed file is marked trashed, keeping its name and category',
   trashed.changed && trashed.docs[0].gone === 'trashed' &&
   trashed.docs[0].name === docA.name && trashed.docs[0].category === docA.category);
var restored = applyFileCheck([Object.assign({}, docA, {gone:'trashed'})], {f_a:{exists:true, trashed:false, name:docA.name}});
ok('a file restored out of the trash loses its gone mark',
   restored.changed && restored.docs[0].gone === undefined);
var renamed = applyFileCheck([docA], {f_a:{exists:true, trashed:false, name:'New name.pdf'}});
ok('a rename in Drive updates the name',
   renamed.changed && renamed.docs[0].name === 'New name.pdf');
var moved = applyFileCheck([docA], {f_a:{exists:true, trashed:false, folderName:'Pricing'}});
ok('a move records the folder it landed in',
   moved.changed && moved.docs[0].folderName === 'Pricing');
var same = applyFileCheck([docA], {f_a:{exists:true, trashed:false, name:docA.name, url:docA.link, folderName:null}});
ok('nothing to report leaves changed false',
   same.changed === false);

var found = [
  {fileId:'f_a', name:'Menu.pdf', link:'https://old-link', category:'Menu'},          // already tracked
  {fileId:'f_new', name:'Banquet menu.pdf', link:'https://new-link', category:'Menu'} // dropped straight into Drive
];
var picked = applyFoundFiles([docA], [], found);
ok('a file Drive has and the console already knows is not duplicated',
   picked.docs.filter(function(d){ return d.fileId === 'f_a'; }).length === 1);
ok('a file uploaded straight into Drive is picked up',
   picked.changed && picked.docs.some(function(d){ return d.fileId === 'f_new' && d.name === 'Banquet menu.pdf'; }));
var removedOne = applyFoundFiles([], ['f_new'], found.filter(function(f){ return f.fileId === 'f_new'; }));
ok('a file someone removed from the list on purpose is not brought back',
   removedOne.changed === false && removedOne.docs.length === 0);
ok('nothing new in Drive leaves changed false',
   applyFoundFiles([docA, docB], [], [{fileId:'f_a'}, {fileId:'f_b'}]).changed === false);
ok('a file with no category is filed as Other',
   applyFoundFiles([], [], [{fileId:'f_c', name:'Loose file.pdf', link:'https://l'}]).docs[0].category === 'Other');

/* The skill's own wording (redaction rules, wording tests, template shapes)
 * moved out of this public repo in "Stop embedding the private skill text in
 * the public repo" and lives only in the private hermetic-venue-config-writer
 * skill now, so it cannot be checked from here any more -- there used to be
 * dozens of ok()s above pinned to exact SKILL_TEXT phrases, and they went with
 * it. What is left in index.html, and so what is left to test, is the one line
 * that tells Claude which skill to run and what to run it on. */
section('What the console tells Claude to run');
ok('it names the skill Claude must invoke',
   SKILL_INVOCATION.indexOf('"hermetic-venue-config-writer" skill') !== -1, SKILL_INVOCATION);
ok('it asks for all three documents',
   SKILL_INVOCATION.indexOf('LOCATION DETAILS') !== -1 &&
   SKILL_INVOCATION.indexOf('CAMPAIGN INSTRUCTIONS') !== -1 &&
   SKILL_INVOCATION.indexOf('OPEN QUESTIONS FOR THE CLIENT') !== -1, SKILL_INVOCATION);
ok('it points the skill at the material that follows, not at nothing',
   SKILL_INVOCATION.indexOf('from the material below') !== -1, SKILL_INVOCATION);

print('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) throw new Error(fail + ' failing');
