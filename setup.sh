#!/bin/zsh
# One-time setup on a new Mac (Apple Silicon). Run from the repo folder:  ./setup.sh
# Installs Ollama + Node, downloads the model, trusts a localhost HTTPS certificate,
# registers the add-in with Excel, then runs ./test.sh.
set -euo pipefail
cd "$(dirname "$0")"
MODEL="${MODEL:-qwen3:4b-instruct}"   # override: MODEL=qwen3:8b ./setup.sh
step() { print -P "\n%B== $1%b"; }

step "1/7 Checking the Mac"
[ "$(uname -s)" = "Darwin" ] || { echo "macOS only."; exit 1; }
[ "$(uname -m)" = "arm64" ] || echo "! Not Apple Silicon — models will run slowly on CPU."
mem_gb=$(( $(sysctl -n hw.memsize) / 1073741824 ))
echo "RAM: ${mem_gb} GB"
[ "$mem_gb" -lt 16 ] && echo "  On ${mem_gb} GB: close big apps (Chrome, Docker) while using LOKUP; the model needs ~3.3 GB."

step "2/7 Homebrew"
if ! command -v brew >/dev/null; then
  [ -x /opt/homebrew/bin/brew ] && eval "$(/opt/homebrew/bin/brew shellenv)"
fi
if ! command -v brew >/dev/null; then
  echo "Homebrew is required. Install it with:"
  echo '  /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
  echo "then run ./setup.sh again."
  exit 1
fi
echo "✓ $(brew --version | head -1)"

step "3/7 Ollama + Node"
for pkg in ollama node; do
  if brew list "$pkg" >/dev/null 2>&1; then echo "✓ $pkg already installed"; else brew install "$pkg"; fi
done
brew services start ollama >/dev/null || true
for i in {1..30}; do curl -s localhost:11434/api/version >/dev/null && break; sleep 1; done
curl -s localhost:11434/api/version >/dev/null && echo "✓ Ollama running (starts automatically at login)" || { echo "✗ Ollama did not start"; exit 1; }

step "4/7 Model: $MODEL (~2.5 GB download)"
if ollama list | awk '{print $1}' | grep -qx "$MODEL"; then echo "✓ already downloaded"; else ollama pull "$MODEL"; fi

step "5/7 HTTPS certificate for https://localhost:3000"
if [ -f ~/.office-addin-dev-certs/localhost.crt ] && security find-certificate -c "Developer CA for Microsoft Office Add-ins" >/dev/null 2>&1; then
  echo "✓ already installed and trusted"
else
  echo "This creates a certificate that only covers 'localhost' and asks for your Mac password to trust it."
  npx --yes office-addin-dev-certs@3.0.1 install
fi

step "6/7 Register the add-in with Excel"
if [ -d "/Applications/Microsoft Excel.app" ]; then
  wef=~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
  mkdir -p "$wef"
  cp addin/manifest.xml "$wef/local-llm.manifest.xml"
  echo "✓ manifest copied to Excel's developer add-in folder"
else
  echo "! Microsoft Excel not found in /Applications — install it, then rerun this step."
fi
chmod +x start.sh test.sh

step "7/7 Quick test"
./test.sh "$MODEL" || echo "! Test reported problems — see above."

cat <<'EOF'

Done. To use it:
  1. ./start.sh                         (keep the window open while using Excel)
  2. Open Excel → Home → Add-ins → Developer Add-ins → Local LLM
  3. In a cell:  =LOKUP.V("Total revenue for North", A1:F13)
                 =LOKUP.H("Classify the sentiment: Positive or Negative", A2)
EOF
