#!/usr/bin/env bash
set -euo pipefail

if [[ $# -eq 0 ]]; then
  echo "Usage: $(basename "$0") -- [arguments for go run .]" >&2
  echo "Example: $(basename "$0") -- -spreadsheet <id> -config-xlsx Schedule.xlsx -lookup-value 'โอเลี้ยง' -values '[[\"โอเลี้ยง\"]]'" >&2
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

SCOPES="https://www.googleapis.com/auth/cloud-platform,https://www.googleapis.com/auth/spreadsheets"

echo "Requesting ADC login with scopes: ${SCOPES}" >&2

gcloud auth application-default login --scopes="${SCOPES}"

if [[ -n "${GCP_QUOTA_PROJECT:-}" ]]; then
  echo "Setting quota project to ${GCP_QUOTA_PROJECT}" >&2
  gcloud auth application-default set-quota-project "${GCP_QUOTA_PROJECT}"
fi

echo "Running go run . $*" >&2

cd "${SCRIPT_DIR}"
exec go run . "$@"
