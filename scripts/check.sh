#!/usr/bin/env bash
# Quick checks the push safety step runs before any push (the repo has no
# real test suite; Apps Script code cannot run outside Google).
# 1. Code.gs parses as JavaScript.  2. The import script parses as Python.
# 3. appsscript.json is valid JSON.  4. All four pages carry one build stamp.
set -euo pipefail
cd "$(dirname "$0")/.."

node -e "new Function(require('fs').readFileSync('Code.gs','utf8'))" \
  && echo "check: Code.gs parses"

python3 -c "import ast,sys; ast.parse(open('import/limesurvey_to_history.py').read())" \
  && echo "check: import/limesurvey_to_history.py parses"

python3 -c "import json; json.load(open('appsscript.json'))" \
  && echo "check: appsscript.json is valid JSON"

stamps=$(grep -ohE "DS-[0-9]{4}-[0-9]{2}-[0-9]{2}-[0-9]+" Code.gs Index.html Board.html Dashboard.html | sort -u)
count=$(printf '%s\n' "$stamps" | grep -c .)
if [ "$count" -ne 1 ]; then
  echo "check FAILED: the four files carry $count different build stamps:" >&2
  printf '  %s\n' $stamps >&2
  exit 1
fi
echo "check: one build stamp across the four files ($stamps)"
