# ──────────────────────────────────────────────────────────────────────────────
# Medmin mcp-o365 installer for Windows (PowerShell)
# Installs the mcp-o365 MCP server and Teams Meeting Analyser skill for
# Claude Code.
#
# Standard usage (run in PowerShell as your normal user — no admin needed):
#   irm https://raw.githubusercontent.com/MedminGroup/mcp-o365/main/install.ps1 | iex
#
# Or with a local build (for development / testing):
#   .\install.ps1 -Source C:\path\to\mcp-O365\dist\index.js
#
# Requirements: Node.js >= 20, Claude Code
# ──────────────────────────────────────────────────────────────────────────────
param(
    [string]$Source = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Configuration ──────────────────────────────────────────────────────────────
$GitHubRepo    = "MedminGroup/mcp-o365"
$AzureClientId = "1cef0b95-5220-4bfa-a2f4-661da5cfcc55"
$AzureTenantId = "389366c7-63a8-42d0-8a1f-1df099d3eec1"
$InstallDir    = Join-Path $HOME ".medmin\mcp-o365"
$PluginName    = "medmin-skills"
$PluginVersion = "1.0.0"
$SkillName     = "teams-meeting-analyser"

# ── Helpers ────────────────────────────────────────────────────────────────────
function Info($msg)    { Write-Host "▶ $msg" -ForegroundColor Cyan }
function Ok($msg)      { Write-Host "✓ $msg" -ForegroundColor Green }
function Warn($msg)    { Write-Host "⚠ $msg" -ForegroundColor Yellow }
function Header($msg)  { Write-Host "`n── $msg ──" -ForegroundColor White }
function Fail($msg)    { Write-Host "✗ $msg" -ForegroundColor Red; exit 1 }

# ── Prerequisites ──────────────────────────────────────────────────────────────
Header "Checking prerequisites"

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    Fail "Node.js is not installed. Download and install Node.js 20 or later from https://nodejs.org then re-run this script."
}
$nodeVersion = [int](node -e "process.stdout.write(process.version.slice(1).split('.')[0])")
if ($nodeVersion -lt 20) {
    Fail "Node.js $(node --version) found, but version 20+ is required. Update from https://nodejs.org"
}
Ok "Node.js $(node --version)"

$claudeJson = Join-Path $HOME ".claude.json"
if (-not (Test-Path $claudeJson)) {
    Fail "Claude Code config not found at $claudeJson — install Claude Code first."
}
Ok "Claude Code config found"

# ── Install MCP server ─────────────────────────────────────────────────────────
Header "Installing MCP server"

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
$DistFile = Join-Path $InstallDir "index.js"

if ($Source -ne "") {
    if (-not (Test-Path $Source)) { Fail "Source file not found: $Source" }
    Info "Copying from $Source"
    Copy-Item $Source $DistFile -Force
} else {
    # Resolve latest release asset URL from GitHub API
    $ReleasesUrl = "https://api.github.com/repos/$GitHubRepo/releases/latest"
    Info "Fetching latest release info..."
    try {
        $release = Invoke-RestMethod -Uri $ReleasesUrl -Headers @{ "User-Agent" = "mcp-o365-installer" }
    } catch {
        Fail "Could not reach GitHub API: $_"
    }
    $asset = $release.assets | Where-Object { $_.name -eq "index.js" } | Select-Object -First 1
    if (-not $asset) { Fail "index.js not found in latest release assets." }
    Info "Downloading $($asset.browser_download_url)"
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $DistFile -UseBasicParsing
}
Ok "MCP server installed to $DistFile"

# ── Configure Claude Code ──────────────────────────────────────────────────────
Header "Configuring Claude Code"

$config = Get-Content $claudeJson -Raw | ConvertFrom-Json

# ConvertFrom-Json returns a PSCustomObject; add mcpServers if missing
if (-not ($config.PSObject.Properties.Name -contains "mcpServers")) {
    $config | Add-Member -NotePropertyName "mcpServers" -NotePropertyValue ([PSCustomObject]@{})
}

$existed = $config.mcpServers.PSObject.Properties.Name -contains "mcp-o365"

$serverEntry = [PSCustomObject]@{
    type    = "stdio"
    command = "node"
    args    = @($DistFile)
    env     = [PSCustomObject]@{
        AZURE_CLIENT_ID = $AzureClientId
        AZURE_TENANT_ID = $AzureTenantId
    }
}

if ($existed) {
    $config.mcpServers."mcp-o365" = $serverEntry
} else {
    $config.mcpServers | Add-Member -NotePropertyName "mcp-o365" -NotePropertyValue $serverEntry
}

$config | ConvertTo-Json -Depth 10 | Set-Content $claudeJson -Encoding UTF8
$action = if ($existed) { "Updated" } else { "Registered" }
Ok "$action mcp-o365 in Claude Code"

# ── Install Teams Meeting Analyser skill ───────────────────────────────────────
Header "Installing Teams Meeting Analyser skill"

$skillDir = Join-Path $HOME ".claude\plugins\cache\local-plugins\$PluginName\$PluginVersion\skills\$SkillName"
New-Item -ItemType Directory -Force -Path $skillDir | Out-Null

$skillContent = @'
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

Save analysis files to the Desktop with the naming convention:
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
'@

Set-Content -Path (Join-Path $skillDir "SKILL.md") -Value $skillContent -Encoding UTF8
Ok "Skill file written"

# Register the plugin
$pluginsJson = Join-Path $HOME ".claude\plugins\installed_plugins.json"
if (Test-Path $pluginsJson) {
    $plugins = Get-Content $pluginsJson -Raw | ConvertFrom-Json
} else {
    $plugins = [PSCustomObject]@{ version = 2; plugins = [PSCustomObject]@{} }
}

$key = "$PluginName@local-plugins"
$now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z")
$installPath = Join-Path $HOME ".claude\plugins\cache\local-plugins\$PluginName\$PluginVersion"

$pluginEntry = @([PSCustomObject]@{
    scope       = "user"
    installPath = $installPath
    version     = $PluginVersion
    installedAt = $now
    lastUpdated = $now
})

if ($plugins.plugins.PSObject.Properties.Name -contains $key) {
    $plugins.plugins.$key = $pluginEntry
} else {
    $plugins.plugins | Add-Member -NotePropertyName $key -NotePropertyValue $pluginEntry
}

$plugins | ConvertTo-Json -Depth 10 | Set-Content $pluginsJson -Encoding UTF8
Ok "Plugin '$key' registered"

# ── Done ───────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "── Installation complete ──" -ForegroundColor Green
Write-Host ""
Write-Host "  Next steps:"
Write-Host "  1. Restart Claude Code (quit and reopen), or run /reload-plugins"
Write-Host "  2. In a new Claude conversation, type: accounts_add"
Write-Host "  3. Open the link shown and sign in with your medmin.co.uk account"
Write-Host "  4. Type: accounts_complete"
Write-Host "  5. You're ready. Try:"
Write-Host "       `"Analyse last week's [meeting name] for [your name]`""
Write-Host ""
