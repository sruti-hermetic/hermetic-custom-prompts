#!/bin/sh
# Runs the tests against apps-script.gs and index.html, without Apps Script
# and without a browser.
#
# Neither file can be run where it lives. apps-script.gs talks to
# SpreadsheetApp, PropertiesService, LockService and ScriptApp, none of which
# exist off Google's servers; apps-script-mock.js stands in for them, with a
# real grid of cells, real merges, real property storage and a trigger list
# with Google's cap on it. index.html is a page, so the sync functions under
# test are cut out of it here and run against a stub localStorage.
#
# Either way the code under test is the code that ships -- nothing is copied.
#
# No install: macOS ships the JavaScript engine this uses.
#
#   sh tests/run.sh
set -e
DIR=$(cd "$(dirname "$0")" && pwd)
JSC=/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc
[ -x "$JSC" ] || { echo "No JavaScriptCore at $JSC (macOS only)."; exit 1; }

echo "=== The readable view (apps-script.gs) ==="
TMP=$(mktemp -t hermetic-view)
{
  echo "load('$DIR/apps-script-mock.js');"
  echo "load('$DIR/../apps-script.gs');"
  sed -e "/^load(/d" "$DIR/readable-view-test.js"
} > "$TMP"
"$JSC" "$TMP"

echo ""
echo "=== Sending work to the Sheet (index.html) ==="
CUT=$(mktemp -t hermetic-cut)
python3 - "$DIR/../index.html" "$CUT" <<'PY'
import re, sys
src = open(sys.argv[1], encoding='utf-8').read()
def grab(name):
    m = re.search(r'\nfunction ' + name + r'\(.*?\n\}\n', src, re.S)
    if not m: raise SystemExit('tests: could not find ' + name + ' in index.html')
    return m.group(0)
parts = [
  re.search(r"const KV_PREFIX = .*?\nconst KV_LOCAL_ONLY = new Set\(\[.*?\]\);", src, re.S).group(0),
  re.search(r"const RECORD_KINDS = \{.*?\n\};", src, re.S).group(0),
]
parts += [grab(n) for n in ['localRead', 'localKeys', 'kindOfKey', 'queueLocalOnly',
                            'getAtPath', 'setAtPath', 'newVenueData', 'serializePart',
                            'foldLegacyBlocks', 'applyFileCheck', 'applyFoundFiles']]
# the schema blocks, and the line that tells Claude which skill to run --
# the skill's own text used to live here too, but it moved out into the
# private skill itself (see "Stop embedding the private skill text in the
# public repo"), so there is nothing left here to check it against.
for pattern in [r"const LOCATION_SCHEMA = \[.*?\n\];",
                r"const CAMPAIGN_SCHEMA = \[.*?\n\];",
                r"const POLICY_FIELDS = \[.*?\n\];",
                r"const FAQ_FIELDS = \[.*?\n\];",
                r"const RETIRED_FIELDS = \[.*?\];",
                r"const FILE_SOURCED_LABEL = .*?;",
                r"const SKILL_INVOCATION = .*?;"]:
    m = re.search(pattern, src, re.S)
    if not m: raise SystemExit('tests: could not find ' + pattern)
    parts.append(m.group(0))
open(sys.argv[2], 'w', encoding='utf-8').write('\n'.join(parts))
PY
TMP2=$(mktemp -t hermetic-client)
sed -e "s|load('UNDER_TEST');.*|load('$CUT');|" "$DIR/client-sync-test.js" > "$TMP2"
"$JSC" "$TMP2"
