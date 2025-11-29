#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "Launching the Google Sheets updater..." >&2

cd "${SCRIPT_DIR}"
exec go run . "$@"
