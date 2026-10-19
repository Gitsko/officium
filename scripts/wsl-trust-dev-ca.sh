#!/usr/bin/env bash
# Trust the office-addin-dev-certs CA (generated inside WSL) in the Windows
# CurrentUser root store, so Outlook/WebView2 on Windows accepts https://localhost:3000.
set -euo pipefail

ca="$HOME/.office-addin-dev-certs/ca.crt"
if [[ ! -f "$ca" ]]; then
  echo "No CA at $ca - run 'npx office-addin-dev-certs install' first." >&2
  exit 1
fi

win_tmp="$(wslpath "$(cmd.exe /c 'echo %TEMP%' 2>/dev/null | tr -d '\r')")"
cp "$ca" "$win_tmp/officium-dev-ca.crt"
certutil.exe -user -addstore Root "$(wslpath -w "$win_tmp/officium-dev-ca.crt")"
