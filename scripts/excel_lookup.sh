#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "Launching the Excel Lookup..." >&2

cd "${SCRIPT_DIR}"
cd ..
TZ=Asia/Bangkok exec  go run ./cmd/lookupscan
