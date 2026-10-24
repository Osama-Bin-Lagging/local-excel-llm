#!/bin/zsh
# Start everything LOKUP needs: Ollama (model server) + the add-in server (https://localhost:3000).
# Keep this window open while you use Excel; Ctrl+C stops the add-in server (Ollama keeps running).
set -u
cd "$(dirname "$0")"

if ! curl -s localhost:11434/api/version >/dev/null; then
  echo "Starting Ollama…"
  if command -v brew >/dev/null && brew list ollama >/dev/null 2>&1; then
    brew services start ollama >/dev/null
  else
    nohup ollama serve >/tmp/ollama.log 2>&1 &
  fi
  for i in {1..30}; do curl -s localhost:11434/api/version >/dev/null && break; sleep 1; done
fi
curl -s localhost:11434/api/version >/dev/null || { echo "✗ Ollama did not start"; exit 1; }
echo "✓ Ollama running on http://localhost:11434"

if curl -s -o /dev/null https://localhost:3000/functions.json; then
  echo "✓ Add-in server already running on https://localhost:3000"
  exit 0
fi
mkdir -p addin/logs
echo "✓ Starting add-in server on https://localhost:3000 (Ctrl+C to stop)"
LOG_PROMPTS="${LOG_PROMPTS:-0}" exec node addin/server.js
