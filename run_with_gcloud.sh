#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

SCOPES="https://www.googleapis.com/auth/cloud-platform,https://www.googleapis.com/auth/spreadsheets"

echo "Requesting ADC login with scopes: ${SCOPES}" >&2

gcloud auth application-default login --scopes="${SCOPES}"

if [[ -n "${GCP_QUOTA_PROJECT:-}" ]]; then
  echo "Setting quota project to ${GCP_QUOTA_PROJECT}" >&2
  gcloud auth application-default set-quota-project "${GCP_QUOTA_PROJECT}"
fi

echo "Launching the interactive Google Sheets updater..." >&2

cd "${SCRIPT_DIR}"
exec go run . "$@"
