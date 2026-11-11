#!/bin/sh
# Installs (or with --uninstall removes) a systemd user service that runs the Outlook MCP server headless.
set -eu

unit_name=outlook-mcp.service
unit_dir="${XDG_CONFIG_HOME:-$HOME/.config}/systemd/user"
unit="$unit_dir/$unit_name"

if [ "${1:-}" = "--uninstall" ]; then
  systemctl --user disable --now "$unit_name" 2>/dev/null || true
  rm -f "$unit"
  systemctl --user daemon-reload
  echo "Removed $unit_name. The browser profile and .env were left in place."
  exit 0
fi

project=$(cd "$(dirname "$0")/.." && pwd)
node=$(command -v node) || { echo "node is not on PATH." >&2; exit 1; }
# Prefer a version manager's stable "latest" link over a versioned path that disappears on upgrade.
case "$node" in
  */mise/installs/node/*/bin/node)
    stable="${node%/node/*}/node/latest/bin/node"
    [ -x "$stable" ] && node=$stable ;;
esac

[ -f "$project/.env" ] || { echo "Create $project/.env with MCP_TOKEN first (see README)." >&2; exit 1; }
grep -q '^MCP_TOKEN=.\{32,\}' "$project/.env" || { echo "$project/.env needs an MCP_TOKEN of at least 32 characters." >&2; exit 1; }
[ -f "$project/dist/server.js" ] || { echo "Run npm run build first." >&2; exit 1; }

if pgrep -f "^[^ ]*node .*dist/server\.js" >/dev/null 2>&1 \
   && ! systemctl --user is-active --quiet "$unit_name"; then
  echo "An Outlook MCP server is already running outside systemd. Stop it first (Ctrl+C in its terminal)." >&2
  exit 1
fi

mkdir -p "$unit_dir"
cat > "$unit" <<UNIT
[Unit]
Description=Outlook Web MCP server
Documentation=file://$project/README.md
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
WorkingDirectory=$project
ExecStart=$node --env-file=$project/.env $project/dist/server.js
# Values set here win over .env. The service is headless; npm run login opens a window when needed.
Environment=OUTLOOK_INTERACTIVE=0
Restart=on-failure
RestartSec=10
TimeoutStopSec=15
UMask=0077
NoNewPrivileges=yes
PrivateTmp=yes
ProtectSystem=full

[Install]
WantedBy=default.target
UNIT

systemctl --user daemon-reload
systemctl --user enable "$unit_name" >/dev/null
systemctl --user restart "$unit_name"
echo "Installed and started $unit_name"
echo "  status:  systemctl --user status outlook-mcp   (or npm run status)"
echo "  logs:    journalctl --user -u outlook-mcp -f"
if [ "$(loginctl show-user "$(id -un)" -p Linger --value 2>/dev/null)" != "yes" ]; then
  echo "  The service runs while you are logged in. To keep it running after logout: loginctl enable-linger"
fi
