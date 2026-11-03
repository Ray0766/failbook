#!/usr/bin/env bash
# The evidence chain for one case, end to end, with no AI and no API key:
#   fresh input -> reproduce the failure -> grade it (evidence) -> the fix's self-check
#   -> a corrected output -> grade again.
# Usage: scripts/demo-evidence-chain.sh [case-id]   (default xlsx-001; needs LibreOffice)
set -euo pipefail
case_id="${1:-xlsx-001}"
work="$(mktemp -d)"
trap 'rm -rf "$work"' EXIT
step() { printf '\n$ %s\n' "$*"; "$@" || true; }

echo "# Evidence chain for $case_id ($(date -u +%Y-%m-%d), failbook $(failbook --version | cut -d' ' -f2))"
step failbook prepare "$case_id" "$work/attempt"
cp -r "$work/attempt" "$work/original"
step failbook solve "$case_id" "$work/attempt" --kind failure
step failbook grade "$case_id" "$work/attempt"
file="$(ls "$work/attempt"/*.xlsx | head -1)"
step failbook audit "$file" --before "$work/original/$(basename "$file")"
rm -rf "$work/fixed" && cp -r "$work/original" "$work/fixed"
step failbook solve "$case_id" "$work/fixed" --kind reference
step failbook grade "$case_id" "$work/fixed"
step failbook audit "$work/fixed/$(basename "$file")" --before "$work/original/$(basename "$file")"
