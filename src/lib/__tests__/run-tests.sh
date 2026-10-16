#!/usr/bin/env bash
# Smoke-test every excel-api.ts action against the fake Office.js host.
set -euo pipefail
cd "$(dirname "$0")/../../.."
./node_modules/.bin/esbuild src/lib/excel-api.ts --bundle --format=esm \
  --platform=browser --outfile=src/lib/__tests__/api.mjs --log-level=error
node src/lib/__tests__/excel-api.test.mjs
