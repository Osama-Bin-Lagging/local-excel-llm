#!/bin/zsh
# Make Excel load the add-in's latest code. Excel for Mac caches add-in files and ignores version bumps.
# Usage: quit Excel, run ./reload.sh, reopen Excel, then Home → Add-ins → insert the add-in(s) again.
set -eu
cd "$(dirname "$0")"
if pgrep -x "Microsoft Excel" >/dev/null; then echo "Quit Excel first (⌘Q), then rerun ./reload.sh"; exit 1; fi
lib=~/Library/Containers/com.microsoft.Excel/Data/Library
old=$(mktemp -d /tmp/excel-addin-cache.XXXXXX)
for d in "$lib/Application Support/Microsoft/Office/16.0/Wef" "$lib/Caches/WebKit"; do
  [ -e "$d" ] && mv "$d" "$old/" || true
done
wef=~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
mkdir -p "$wef"
cp addin/manifest.xml "$wef/local-llm.manifest.xml"
cp addin/manifest-workbook.xml "$wef/local-llm-workbook.manifest.xml"
echo "✓ Add-in cache cleared (old copy in $old). Open Excel → Home → Add-ins → insert Local LLM / Local LLM Workbook."
