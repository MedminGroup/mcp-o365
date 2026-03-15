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
# Installs Node.js and Claude Code CLI automatically if missing (no admin required).
# ──────────────────────────────────────────────────────────────────────────────
set -euo pipefail

GITHUB_REPO="MedminGroup/mcp-o365"
AZURE_CLIENT_ID="1cef0b95-5220-4bfa-a2f4-661da5cfcc55"
AZURE_TENANT_ID="389366c7-63a8-42d0-8a1f-1df099d3eec1"
INSTALL_DIR="$HOME/.medmin/mcp-o365"
PLUGIN_NAME="medmin-skills"
PLUGIN_VERSION="1.0.0"
SKILL_NAME="teams-meeting-analyser"
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
  # Source nvm in the current script session
  [ -s "$NVM_DIR/nvm.sh" ] && \. "$NVM_DIR/nvm.sh"
  nvm install 22 2>&1 | grep -E 'Now using|Downloading|installed' || true
  nvm use 22 >/dev/null 2>&1
  ok "Node.js $(node --version) installed via nvm"
}

# Source nvm if it's already installed but not on PATH
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

# ── Claude Code CLI ───────────────────────────────────────────────────────────
header "Checking Claude Code CLI"

if ! command -v claude &>/dev/null; then
  info "Installing Claude Code CLI (this may take a minute)..."
  npm install -g @anthropic-ai/claude-code
  command -v claude &>/dev/null || error "claude command not found after npm install."
  ok "Claude Code CLI installed: $(claude --version 2>/dev/null | head -1)"
else
  ok "Claude Code CLI: $(claude --version 2>/dev/null | head -1)"
fi

# Config is created on first launch — create a minimal one if not yet present
CLAUDE_JSON="$HOME/.claude.json"
if [[ ! -f "$CLAUDE_JSON" ]]; then
  info "Config not yet created — writing minimal config at $CLAUDE_JSON"
  echo '{"mcpServers":{}}' > "$CLAUDE_JSON"
  ok "Created $CLAUDE_JSON"
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

# ── Configure Claude Code ──────────────────────────────────────────────────────
header "Configuring Claude Code"

python3 - "$CLAUDE_JSON" "$DIST_FILE" "$AZURE_CLIENT_ID" "$AZURE_TENANT_ID" <<'PYEOF'
import json, sys, os
claude_json, dist_file, client_id, tenant_id = sys.argv[1:]
with open(claude_json) as f:
    config = json.load(f)
config.setdefault("mcpServers", {})
existed = "mcp-o365" in config["mcpServers"]
config["mcpServers"]["mcp-o365"] = {
    "type": "stdio", "command": "node", "args": [dist_file],
    "env": {"AZURE_CLIENT_ID": client_id, "AZURE_TENANT_ID": tenant_id}
}
with open(claude_json, "w") as f:
    json.dump(config, f, indent=2)
print(("Updated" if existed else "Registered") + " mcp-o365.")
PYEOF

ok "mcp-o365 registered in Claude Code"

# ── Install skills ─────────────────────────────────────────────────────────────
header "Installing Medmin skills"

# ── Skill 1: Teams Meeting Analyser ───────────────────────────────────────────
SKILL_DIR="$HOME/.claude/plugins/cache/local-plugins/$PLUGIN_NAME/$PLUGIN_VERSION/skills/$SKILL_NAME"
mkdir -p "$SKILL_DIR"

cat > "$SKILL_DIR/SKILL.md" << 'SKILL_EOF'
---
name: teams-meeting-analyser
description: Fetches Microsoft Teams meeting transcripts via the mcp-o365 MCP server and analyses communication patterns, speaking ratios, filler words, conflict avoidance, facilitation style, and key decisions for any named participant. Use when asked to analyse a Teams meeting, a meeting transcript, or a person's communication style in a recorded meeting.
---

# Teams Meeting Analyser

Fetches live transcripts from Microsoft Teams via the Microsoft Graph API and produces
deep communication pattern analysis for one or more named participants.

## When to Use This Skill

- "Analyse [person]'s contribution to [meeting name]"
- "What was decided in last Thursday's [meeting]?"
- "How did [person] communicate in the [meeting] meeting?"
- "Pull the transcript from [meeting name] and analyse it"
- "Analyse this week's / last week's [recurring meeting name]"

## Prerequisites

- Microsoft 365 account authenticated via the mcp-o365 MCP server
  (run `accounts_add` in Claude if not yet signed in)
- Meeting must have had transcription enabled (Teams: ... → Start transcription)

## What This Skill Does

1. Resolves the meeting from the user's calendar using `calendar_list_events`
2. Retrieves the Teams online meeting object via `meetings_get_by_join_url`
3. Lists available transcripts via `meetings_list_transcripts`
4. Downloads the VTT transcript via `meetings_get_transcript`
5. Analyses communication patterns for named participants
6. Produces a structured Meeting Insights Summary

## Step-by-Step Workflow

### 0. Confirm signed-in account

Call `accounts_list` to get the user's account identifier. Use that value as the
`account` parameter in all subsequent calls.

### 1. Find the calendar event

Use `calendar_list_events` to locate the meeting. Ask the user for the date if not
provided. Search a +/-1 day window to allow for timezone differences.

    calendar_list_events(
      account="<from accounts_list>",
      start="<date>T00:00:00",
      end="<date>T23:59:59"
    )

Note the event's `joinWebUrl` from the response. If the event has no `joinWebUrl`
it was not a Teams meeting and has no transcript.

### 2. Get the online meeting ID

    meetings_get_by_join_url(
      join_url="<joinWebUrl from step 1>",
      account="<account>"
    )

Note the `id` field from the response.

### 3. List transcripts

    meetings_list_transcripts(
      meeting_id="<id from step 2>",
      account="<account>"
    )

For recurring meetings, multiple transcripts will be listed. Match by
`createdDateTime` to find the correct occurrence. Note the target transcript's `id`.

### 4. Download the VTT transcript

    meetings_get_transcript(
      meeting_id="<id from step 2>",
      transcript_id="<id from step 3>",
      account="<account>"
    )

The response is raw VTT text with speaker labels and timestamps.

### 5. Validate the transcript

Before analysing, extract basic stats from the VTT:

- All unique speaker names (from <v ...> tags)
- Turn count per speaker
- Approximate duration (last timestamp minus first)

Confirm with the user who to focus on if not already specified.

### 6. Analyse the transcript

For each participant to be analysed, produce a Meeting Insights Summary covering:

**Speaking Statistics**
- Turn count and percentage per speaker
- Word count and percentage per speaker
- Average words per turn
- Filler word counts: "um", "uh", "like", "you know", "I think", "sort of", "kind of"
- Question vs statement ratio

**Communication Patterns**
- Directness — are statements assertive or hedged?
- Conflict avoidance — hedging language, subject changes, indirect phrasing
- Active listening — building on others' points, clarifying questions, paraphrasing
- Leadership/facilitation — agenda control, drawing out quieter voices, handling disagreement

**Specific Examples**
For each pattern, include:
- Timestamp
- Verbatim quote
- Why it matters
- Better approach (for growth areas)

**Key Decisions and Action Items**
Extract all decisions and action items with owners.

**Strengths and Growth Opportunities**
Minimum 3 strengths, minimum 4 growth opportunities, all with timestamps.

## Output Format

# Meeting Insights Summary — [Name]
**Meeting:** [Name] | [Date] | [Duration]
**Participants:** [List]

## Speaking Ratios
[Table]

## Communication Patterns
[Sections per pattern with quotes and timestamps]

## Key Decisions and Action Items
[Table with owner]

## Strengths
[Numbered list with evidence]

## Growth Opportunities
[Numbered list with timestamps and better alternatives]

## Summary
[2-3 sentence overall assessment]

## Saving Outputs

Save analysis files to ~/Downloads/ with the naming convention:
  meeting-analysis-[firstname]-YYYY-MM-DD.txt

## Known Gotchas

- Recurring meetings share a single online meeting ID. Use `meetings_list_transcripts`
  and filter by `createdDateTime` to find the right occurrence.
- Teams only generates a transcript if "Start transcription" was active during the
  meeting. If the transcripts list is empty, no transcript was recorded.
- `meetings_get_by_join_url` only works for meetings organised by the signed-in account.
  If the user was an attendee (not the organiser), the lookup will return no results.
- If `meetings_get_transcript` returns a 403, contact your administrator.

## Example Prompts

"Analyse last week's Weekly Medmin meeting for Sarah"
"Pull the transcript from the board show and tell and tell me what was decided"
"How did James communicate in Thursday's team meeting?"
"Analyse everyone's contribution to this week's Monday standup"
SKILL_EOF

ok "Teams Meeting Analyser skill installed"

# ── Skill 2: Medmin Guide ──────────────────────────────────────────────────────
GUIDE_DIR="$HOME/.claude/plugins/cache/local-plugins/$PLUGIN_NAME/$PLUGIN_VERSION/skills/medmin-guide"
mkdir -p "$GUIDE_DIR"

cat > "$GUIDE_DIR/SKILL.md" << 'GUIDE_EOF'
---
name: medmin-guide
description: Shows the user a friendly guide to all Medmin tools connected to Claude — what's available, how to use it, and example prompts. Use when the user asks what Claude can do, how the meeting analyser works, what tools are connected, or asks for help getting started.
---

# Medmin Guide

When this skill is invoked, present the following guide to the user exactly as formatted below. Do not summarise it — show it in full.

---

## What Claude can do for you

Claude is connected to your **Microsoft 365** account and can read your calendar, emails, files, and Teams meeting transcripts — all processed privately on your device.

---

### ✅ Teams Meeting Analyser

Fetches live transcripts from your recorded Teams meetings and produces a detailed communication analysis for any participant.

**Before you start — check you're signed in:**
> Ask Claude: *"am I signed in to Microsoft 365?"*

If not signed in yet, run these three steps:
1. Ask Claude: **accounts_add**
2. Open the link shown and sign in with your **medmin.co.uk** account
3. Ask Claude: **accounts_complete**

**How to start a meeting analysis:**

| What you want | What to ask Claude |
|---|---|
| Analyse one person's communication | "Analyse last week's [meeting name] for [name]" |
| Find out what was decided | "What was decided in Thursday's [meeting name]?" |
| Analyse everyone's contribution | "Analyse everyone in this week's [meeting name]" |
| Check someone's communication style | "How did [name] communicate in yesterday's meeting?" |

**What you get back:**
- Speaking ratios and word counts per participant
- Communication patterns: directness, conflict avoidance, active listening, facilitation
- Verbatim quotes with timestamps and coaching suggestions
- Key decisions and action items with owners
- Strengths and growth opportunities

**Requirements:**
- The meeting must have had **Start transcription** active during the call
  (In Teams: click ... -> Start transcription before your meeting begins)
- You must have **organised** the meeting — attendee-only access does not work

**Example prompts to try:**
"Analyse last Monday's weekly standup for Sarah"
"What action items came out of Thursday's board meeting?"
"How did James communicate in the product review — was he direct?"
"Analyse everyone's contribution to this week's all-hands"

Results are saved to your Downloads folder as: meeting-analysis-[firstname]-YYYY-MM-DD.txt

---

### 🔜 HubSpot Integration (coming soon)

Once connected, Claude will be able to:
- Save meeting analysis directly to a contact's HubSpot record
- Pull up a contact's history before a meeting
- Log notes to a contact from a conversation

Example prompts once connected:
"Analyse Tuesday's meeting with John Smith and save it to his HubSpot record"
"What do we know about Sarah Jones before my call with her tomorrow?"
"Log a note on the Acme account — we agreed to push the demo to next week"

---

### Tips

- **Recurring meetings** — just say "last Tuesday's" or "the 10th March" version
- **Multiple accounts** — tell Claude which one: "use my medmin.co.uk account"
- **No transcript found?** — transcription must have been started during the meeting
- **Re-run this guide** — just ask: "show me the Medmin guide" or "what can Claude do?"
GUIDE_EOF

ok "Medmin Guide skill installed"

python3 - "$HOME/.claude/plugins/installed_plugins.json" "$PLUGIN_NAME" "$PLUGIN_VERSION" \
         "$HOME/.claude/plugins/cache/local-plugins/$PLUGIN_NAME/$PLUGIN_VERSION" <<'PYEOF'
import json, sys, os, datetime
plugins_file, name, version, install_path = sys.argv[1:]
plugins = json.load(open(plugins_file)) if os.path.exists(plugins_file) else {"version": 2, "plugins": {}}
key = f"{name}@local-plugins"
now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
plugins["plugins"][key] = [{"scope":"user","installPath":install_path,"version":version,"installedAt":now,"lastUpdated":now}]
json.dump(plugins, open(plugins_file,"w"), indent=2)
print(f"Plugin '{key}' registered.")
PYEOF

ok "Teams Meeting Analyser skill installed"

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
ok "All done! Restart Claude Code and you're ready."
echo ""
