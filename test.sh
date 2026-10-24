#!/bin/zsh
# Quick check that everything works (~1–2 min). Usage: ./test.sh [model]
set -u
cd "$(dirname "$0")"
MODEL="${1:-}"
fail=0

echo "== Ollama"
if curl -s localhost:11434/api/version >/dev/null; then echo "✓ running"; else echo "✗ not running — run ./start.sh"; exit 1; fi

echo "== Model"
want="${MODEL:-$(grep -o 'const DEFAULT_MODEL = "[^"]*"' addin/public/functions.js | cut -d'"' -f2)}"
if ollama list | awk '{print $1}' | grep -qx "$want"; then echo "✓ $want downloaded"; else echo "✗ $want missing — run: ollama pull $want"; exit 1; fi

echo "== HTTPS certificate + add-in server"
started=0
if ! curl -s -o /dev/null https://localhost:3000/functions.json; then
  node addin/server.js >/tmp/lokup-test-server.log 2>&1 &
  srv=$!; started=1; sleep 2
fi
if curl -s -o /dev/null -w "%{http_code}" https://localhost:3000/functions.json | grep -q 200; then
  echo "✓ https://localhost:3000 serves the add-in and the certificate is trusted"
else
  echo "✗ add-in server unreachable or certificate not trusted (see /tmp/lokup-test-server.log; rerun ./setup.sh)"; fail=1
fi
code=$(curl -s -o /dev/null -w "%{http_code}" https://localhost:3000/api/version)
[ "$code" = "200" ] && echo "✓ server → Ollama proxy works" || { echo "✗ proxy to Ollama failed ($code)"; fail=1; }
[ $started = 1 ] && kill $srv 2>/dev/null

echo "== Excel add-in registration"
if [ -f ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/local-llm.manifest.xml ]; then echo "✓ manifest installed for Excel"; else echo "! manifest not installed (is Excel installed? rerun ./setup.sh)"; fi

echo "== LOKUP.H / LOKUP.V on 7 tasks"
node harness/quicktest.js ${MODEL:+$MODEL} || fail=1

exit $fail
