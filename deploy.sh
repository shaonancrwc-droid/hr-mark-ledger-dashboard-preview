#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
wrangler pages deploy public --project-name hr-mark-ledger-dashboard --branch main --commit-dirty=true
