#!/bin/sh
# Runs the readable-view tests against apps-script.gs, without Apps Script.
#
# apps-script.gs cannot be run locally -- it talks to SpreadsheetApp,
# PropertiesService, LockService and ScriptApp, none of which exist off Google's
# servers. apps-script-mock.js stands in for them: a real grid of cells, real
# merges, real property storage, a trigger list with Google's cap on it. The
# tests then drive the actual file, so what passes here is the code that gets
# pasted into the editor.
#
# No install: macOS ships the JavaScript engine this uses.
#
#   sh tests/run.sh
set -e
DIR=$(cd "$(dirname "$0")" && pwd)
JSC=/System/Library/Frameworks/JavaScriptCore.framework/Versions/A/Helpers/jsc
[ -x "$JSC" ] || { echo "No JavaScriptCore at $JSC (macOS only)."; exit 1; }
TMP=$(mktemp -t hermetic-tests)
{
  echo "load('$DIR/apps-script-mock.js');"
  echo "load('$DIR/../apps-script.gs');"
  sed -e "/^load(/d" "$DIR/readable-view-test.js"
} > "$TMP"
"$JSC" "$TMP"
