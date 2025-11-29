#!/usr/bin/env bash
set -euo pipefail
GCP_QUOTA_PROJECT="solid-arcadia-479711-u1"
SCOPES="https://www.googleapis.com/auth/cloud-platform,https://www.googleapis.com/auth/spreadsheets"

if ! command -v gcloud >/dev/null 2>&1; then
  echo "gcloud CLI not found; install the Google Cloud SDK first." >&2
  exit 1
fi

cat <<MSG >&2
Requesting ADC login with scopes:
  ${SCOPES}
MSG

gcloud auth application-default login --scopes="${SCOPES}"

if [[ -n "${GCP_QUOTA_PROJECT:-}" ]]; then
  echo "Setting quota project to ${GCP_QUOTA_PROJECT}" >&2
  gcloud auth application-default set-quota-project "${GCP_QUOTA_PROJECT}"
fi

echo "Authentication complete. Run ./run_with_gcloud.sh to launch the updater." >&2
