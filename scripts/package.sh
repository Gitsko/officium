#!/usr/bin/env bash
# Build the app package (manifest.json plus the icons it references) for manual
# upload. Archive paths must match the paths in the manifest.
set -euo pipefail

cd "$(dirname "$0")/.."
mkdir -p dist
python3 - <<'PY'
import json, zipfile

m = json.load(open("manifest.json"))
files = ["manifest.json", m["icons"]["color"], m["icons"]["outline"]]
with zipfile.ZipFile("dist/officium.zip", "w", zipfile.ZIP_DEFLATED) as z:
    for f in files:
        z.write(f, f)
PY
echo "dist/officium.zip"
