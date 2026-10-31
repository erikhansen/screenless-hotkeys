#!/bin/bash
# Gets the latest version from GitHub and installs it.
set -euo pipefail
cd "$(dirname "$0")"
if ! git pull --ff-only; then
  echo "Update failed while downloading."; /usr/bin/say "Update failed while downloading." || true
  exit 1
fi
./install.sh
