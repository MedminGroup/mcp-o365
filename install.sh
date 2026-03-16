#!/usr/bin/env bash
# ──────────────────────────────────────────────────────────────────────────────
# Medmin mcp-o365 installer (Mac / Linux)
#
# Standard usage — paste into Terminal and press Enter:
#   curl -fsSL https://raw.githubusercontent.com/MedminGroup/mcp-o365/main/install.sh | bash
#
# Dev / testing (local build):
#   bash install.sh --source /path/to/mcp-O365/dist/index.js
#
# Installs Node.js automatically if missing (no admin required).
# Configures the Claude desktop app — no Claude Code CLI needed.
# ──────────────────────────────────────────────────────────────────────────────
set -euo pipefail

GITHUB_REPO="MedminGroup/mcp-o365"
AZURE_CLIENT_ID="1cef0b95-5220-4bfa-a2f4-661da5cfcc55"
AZURE_TENANT_ID="389366c7-63a8-42d0-8a1f-1df099d3eec1"
INSTALL_DIR="$HOME/.medmin/mcp-o365"
NVM_VERSION="0.40.3"

# ── Colours ────────────────────────────────────────────────────────────────────
if [ -t 1 ]; then
  GRN='\033[0;32m' YLW='\033[1;33m' RED='\033[0;31m' BLD='\033[1m' NC='\033[0m'
else
  GRN='' YLW='' RED='' BLD='' NC=''
fi
info()   { echo -e "${GRN}▶${NC} $*"; }
warn()   { echo -e "${YLW}⚠${NC}  $*"; }
error()  { echo -e "${RED}✗${NC}  $*" >&2; exit 1; }
ok()     { echo -e "${GRN}✓${NC} $*"; }
header() { echo -e "\n${BLD}── $* ──${NC}"; }

# ── Arguments ──────────────────────────────────────────────────────────────────
LOCAL_SOURCE=""
while [[ $# -gt 0 ]]; do
  case $1 in
    --source) LOCAL_SOURCE="$2"; shift 2 ;;
    *) warn "Unknown argument: $1"; shift ;;
  esac
done

# ── Node.js ────────────────────────────────────────────────────────────────────
header "Checking Node.js"

install_node_via_nvm() {
  info "Installing Node.js via nvm (no admin required)..."
  export NVM_DIR="$HOME/.nvm"
  curl -fsSL "https://raw.githubusercontent.com/nvm-sh/nvm/v${NVM_VERSION}/install.sh" | bash
  [ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"
  nvm install 22 2>&1 | grep -E 'Now using|Downloading|installed' || true
  nvm use 22 >/dev/null 2>&1
  ok "Node.js $(node --version) installed via nvm"
}

# Source nvm if already installed but not on PATH
export NVM_DIR="$HOME/.nvm"
[ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"

if ! command -v node &>/dev/null; then
  install_node_via_nvm
elif [[ "$(node -e "process.stdout.write(process.version.slice(1).split('.')[0])")" -lt 20 ]]; then
  warn "Node.js $(node --version) is too old (need 20+). Installing a newer version..."
  install_node_via_nvm
else
  ok "Node.js $(node --version)"
fi

# ── Download / copy MCP server ─────────────────────────────────────────────────
header "Installing MCP server"

mkdir -p "$INSTALL_DIR"
DIST_FILE="$INSTALL_DIR/index.js"

if [[ -n "$LOCAL_SOURCE" ]]; then
  [[ -f "$LOCAL_SOURCE" ]] || error "Source file not found: $LOCAL_SOURCE"
  info "Copying from $LOCAL_SOURCE"
  cp "$LOCAL_SOURCE" "$DIST_FILE"
else
  RELEASES_API="https://api.github.com/repos/$GITHUB_REPO/releases/latest"
  info "Fetching latest release..."
  RELEASE_JSON=$(curl -fsSL "$RELEASES_API")
  DIST_URL=$(echo "$RELEASE_JSON" | python3 -c "
import json,sys
assets = json.load(sys.stdin).get('assets', [])
match = next((a['browser_download_url'] for a in assets if a['name'] == 'index.js'), None)
if not match: raise SystemExit('index.js not found in latest release')
print(match)
")
  info "Downloading $DIST_URL"
  curl -fsSL "$DIST_URL" -o "$DIST_FILE"
fi

ok "MCP server installed to $DIST_FILE"

# ── Configure Claude desktop app ───────────────────────────────────────────────
header "Configuring Claude desktop app"

# Locate the Claude desktop app config directory
if [[ "$OSTYPE" == "darwin"* ]]; then
  CLAUDE_CONFIG_DIR="$HOME/Library/Application Support/Claude"
else
  CLAUDE_CONFIG_DIR="$HOME/.config/Claude"
fi

CLAUDE_CONFIG="$CLAUDE_CONFIG_DIR/claude_desktop_config.json"
mkdir -p "$CLAUDE_CONFIG_DIR"

# Create a minimal config if it doesn't exist
if [[ ! -f "$CLAUDE_CONFIG" ]]; then
  info "Creating new config at $CLAUDE_CONFIG"
  echo '{"mcpServers":{}}' > "$CLAUDE_CONFIG"
fi

# Resolve full path to node so Claude desktop app doesn't need node on its PATH
NODE_BIN="$(command -v node)"

python3 - "$CLAUDE_CONFIG" "$DIST_FILE" "$AZURE_CLIENT_ID" "$AZURE_TENANT_ID" "$NODE_BIN" <<'PYEOF'
import json, sys
config_path, dist_file, client_id, tenant_id, node_bin = sys.argv[1:]
with open(config_path) as f:
    config = json.load(f)
config.setdefault("mcpServers", {})
existed = "mcp-o365" in config["mcpServers"]
config["mcpServers"]["mcp-o365"] = {
    "command": node_bin,
    "args": [dist_file],
    "env": {"AZURE_CLIENT_ID": client_id, "AZURE_TENANT_ID": tenant_id}
}
with open(config_path, "w") as f:
    json.dump(config, f, indent=2)
print(("Updated" if existed else "Registered") + " mcp-o365.")
PYEOF

ok "mcp-o365 registered in Claude desktop app"

# ── Install CLAUDE.md ──────────────────────────────────────────────────────────
header "Installing Claude instructions"

cat > "$HOME/CLAUDE.md" << 'CLAUDEMD_EOF'
# Medmin Microsoft 365 — Claude Instructions

You have the **mcp-o365** MCP server connected. It gives you live access to this
user's Microsoft 365 account via the Microsoft Graph API.

## Critical rules — read before every Microsoft 365 task

1. **Always use mcp-o365 tools.** Never use any built-in Microsoft plugin or
   integration. The built-in Microsoft plugin does not work on Medmin accounts.
   mcp-o365 is the only authorised integration.

2. **You have a valid token.** The user is already signed in. Do not ask them to
   sign in unless `accounts_list` returns an empty list.

3. **For any Teams meeting or transcript request**, always follow this sequence:
   - `accounts_list` → confirm the account
   - `calendar_list_events` → find the meeting and its joinWebUrl
   - `meetings_get_by_join_url` → get the meeting ID
   - `meetings_list_transcripts` → list available transcripts
   - `meetings_get_transcript` → download the VTT content
   Then analyse the transcript.

4. **Never guess or fabricate transcript content.** Always fetch it with the tools.

5. **If a tool call fails**, report the exact error. Do not fall back to the
   built-in Microsoft plugin.
CLAUDEMD_EOF

ok "CLAUDE.md installed to $HOME/CLAUDE.md"

# ── Sign-in wizard ─────────────────────────────────────────────────────────────
header "Signing in to Microsoft 365"
echo ""
echo "  A browser window will open — sign in with your medmin.co.uk account."
echo "  If it doesn't open automatically, use the URL printed below."
echo ""

AZURE_CLIENT_ID="$AZURE_CLIENT_ID" \
AZURE_TENANT_ID="$AZURE_TENANT_ID" \
  node "$DIST_FILE" --setup

echo ""
ok "All done! Restart the Claude app and you're ready."
echo "  Open Claude and ask: \"what can you do?\""
echo ""
