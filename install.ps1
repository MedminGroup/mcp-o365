# ──────────────────────────────────────────────────────────────────────────────
# Medmin mcp-o365 installer (Windows PowerShell)
#
# Standard usage — paste into PowerShell and press Enter:
#   irm https://raw.githubusercontent.com/MedminGroup/mcp-o365/main/install.ps1 | iex
#
# Dev / testing (local build):
#   .\install.ps1 -Source C:\path\to\mcp-O365\dist\index.js
#
# Installs Node.js automatically if missing (no admin required).
# Requires: Claude Code
# ──────────────────────────────────────────────────────────────────────────────
param([string]$Source = "")

# ── Logging — captures everything to a file ────────────────────────────────────
$LogFile = Join-Path $env:TEMP "medmin-install-log.txt"
Start-Transcript -Path $LogFile -Force | Out-Null
Write-Host "  [log] Full log: $LogFile" -ForegroundColor DarkGray

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Wrap entire script in try/catch so errors are always visible before exit
trap {
    Write-Host ""
    Write-Host "  ================================================================" -ForegroundColor Red
    Write-Host "  INSTALL FAILED" -ForegroundColor Red
    Write-Host "  ================================================================" -ForegroundColor Red
    Write-Host "  Error: $_" -ForegroundColor Red
    Write-Host "  At:    $($_.InvocationInfo.PositionMessage)" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Full log saved to: $LogFile" -ForegroundColor Yellow
    Write-Host "  Send that file to your administrator." -ForegroundColor Yellow
    Write-Host "  ================================================================" -ForegroundColor Red
    Write-Host ""
    Stop-Transcript | Out-Null
    Read-Host "  Press Enter to close"
    break
}

$GitHubRepo    = "MedminGroup/mcp-o365"
$AzureClientId = "1cef0b95-5220-4bfa-a2f4-661da5cfcc55"
$AzureTenantId = "389366c7-63a8-42d0-8a1f-1df099d3eec1"
$InstallDir    = Join-Path $HOME ".medmin\mcp-o365"
$PluginName    = "medmin-skills"
$PluginVersion = "1.0.0"
$SkillName     = "teams-meeting-analyser"
$NodeDir       = Join-Path $env:LOCALAPPDATA "Programs\nodejs"

function Info($m)   { Write-Host "  [>] $m" -ForegroundColor Cyan }
function Ok($m)     { Write-Host "  [+] $m" -ForegroundColor Green }
function Warn($m)   { Write-Host "  [!] $m" -ForegroundColor Yellow }
function Header($m) { Write-Host "`n  -- $m --" -ForegroundColor White }
function Fail($m)   { Write-Host "  [x] $m" -ForegroundColor Red; exit 1 }

# ── Node.js ────────────────────────────────────────────────────────────────────
Header "Checking Node.js"

function Install-Node {
    Info "Installing Node.js (no admin required)..."

    # Get latest LTS version from nodejs.org
    try {
        $releases = Invoke-RestMethod "https://nodejs.org/dist/index.json" `
            -Headers @{"User-Agent"="mcp-o365-installer"}
        $lts     = ($releases | Where-Object { $_.lts -ne $false } | Select-Object -First 1)
        $version = $lts.version   # e.g. "v22.13.1"
    } catch {
        Fail "Could not fetch Node.js version list: $_"
    }

    $arch    = if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" }
    $zipUrl  = "https://nodejs.org/dist/$version/node-$version-win-$arch.zip"
    $zipFile = Join-Path $env:TEMP "nodejs-install.zip"
    $extract = Join-Path $env:TEMP "nodejs-extract"

    Info "Downloading Node.js $version..."
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing

    Info "Extracting..."
    if (Test-Path $extract) { Remove-Item $extract -Recurse -Force }
    Expand-Archive -Path $zipFile -DestinationPath $extract -Force

    $extracted = Get-ChildItem $extract -Directory | Select-Object -First 1
    if (Test-Path $NodeDir) { Remove-Item $NodeDir -Recurse -Force }
    New-Item -ItemType Directory -Force -Path (Split-Path $NodeDir) | Out-Null
    Move-Item $extracted.FullName $NodeDir

    Remove-Item $zipFile -Force
    Remove-Item $extract -Recurse -Force -ErrorAction SilentlyContinue

    # Add to user PATH (permanent + current session)
    $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    if ($null -eq $userPath -or $userPath -notlike "*$NodeDir*") {
        [Environment]::SetEnvironmentVariable("PATH", "$NodeDir;$userPath", "User")
    }
    $env:PATH = "$NodeDir;$env:PATH"

    Ok "Node.js $version installed to $NodeDir"
}

# Add NodeDir to PATH for this session if it exists but isn't on PATH yet
if ((Test-Path "$NodeDir\node.exe") -and ($env:PATH -notlike "*$NodeDir*")) {
    $env:PATH = "$NodeDir;$env:PATH"
}

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    Install-Node
} else {
    $nodeMajor = [int](node -e "process.stdout.write(process.version.slice(1).split('.')[0])")
    if ($nodeMajor -lt 20) {
        Warn "Node.js $(node --version) is too old (need 20+). Installing a newer version..."
        Install-Node
    } else {
        Ok "Node.js $(node --version)"
    }
}

# Verify node is actually reachable after any install
Info "Verifying node is on PATH..."
$nodeExe = Get-Command node -ErrorAction SilentlyContinue
if (-not $nodeExe) {
    throw "node.exe not found on PATH after install. PATH is: $env:PATH"
}
Ok "node is at $($nodeExe.Source) — $(node --version)"

# ── Claude Code ────────────────────────────────────────────────────────────────
$claudeJson = Join-Path $HOME ".claude.json"
if (-not (Test-Path $claudeJson)) {
    Fail "Claude Code not found ($claudeJson missing). Install Claude Code first: https://claude.ai/download"
}

# ── Download / copy MCP server ─────────────────────────────────────────────────
Header "Installing MCP server"

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
$DistFile = Join-Path $InstallDir "index.js"

if ($Source -ne "") {
    if (-not (Test-Path $Source)) { Fail "Source file not found: $Source" }
    Info "Copying from $Source"
    Copy-Item $Source $DistFile -Force
} else {
    $ReleasesUrl = "https://api.github.com/repos/$GitHubRepo/releases/latest"
    Info "Fetching latest release..."
    try {
        $release = Invoke-RestMethod -Uri $ReleasesUrl `
            -Headers @{"User-Agent"="mcp-o365-installer"}
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

$config  = Get-Content $claudeJson -Raw | ConvertFrom-Json
$existed = $false

if (-not ($config.PSObject.Properties.Name -contains "mcpServers")) {
    $config | Add-Member -NotePropertyName "mcpServers" -NotePropertyValue ([PSCustomObject]@{})
}
if ($config.mcpServers.PSObject.Properties.Name -contains "mcp-o365") { $existed = $true }

$entry = [PSCustomObject]@{
    type    = "stdio"
    command = "node"
    args    = @($DistFile)
    env     = [PSCustomObject]@{ AZURE_CLIENT_ID = $AzureClientId; AZURE_TENANT_ID = $AzureTenantId }
}
if ($existed) { $config.mcpServers."mcp-o365" = $entry }
else          { $config.mcpServers | Add-Member -NotePropertyName "mcp-o365" -NotePropertyValue $entry }

$config | ConvertTo-Json -Depth 10 | Set-Content $claudeJson -Encoding UTF8
Ok ($(if ($existed) { "Updated" } else { "Registered" }) + " mcp-o365 in Claude Code")

# ── Install skill ──────────────────────────────────────────────────────────────
Header "Installing Teams Meeting Analyser skill"

$skillDir = Join-Path $HOME ".claude\plugins\cache\local-plugins\$PluginName\$PluginVersion\skills\$SkillName"
New-Item -ItemType Directory -Force -Path $skillDir | Out-Null

# Write skill file (no backtick-heavy heredoc needed — use Set-Content)
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
  (run accounts_add in Claude if not yet signed in)
- Meeting must have had transcription enabled (Teams: ... -> Start transcription)

## Step-by-Step Workflow

### 0. Confirm signed-in account
Call accounts_list to get the user's account. Use it as the account parameter in all calls.

### 1. Find the calendar event
    calendar_list_events(account="<from accounts_list>", start="<date>T00:00:00", end="<date>T23:59:59")
Note the event's joinWebUrl.

### 2. Get the online meeting ID
    meetings_get_by_join_url(join_url="<joinWebUrl>", account="<account>")
Note the id field.

### 3. List transcripts
    meetings_list_transcripts(meeting_id="<id>", account="<account>")
Match by createdDateTime for recurring meetings.

### 4. Download the transcript
    meetings_get_transcript(meeting_id="<id>", transcript_id="<id>", account="<account>")

### 5. Analyse the transcript
Speaking Statistics, Communication Patterns, Key Decisions, Strengths, Growth Opportunities.

## Output Format
# Meeting Insights Summary - [Name]
**Meeting:** [Name] | [Date] | [Duration]
[Full analysis sections]

## Saving Outputs
Save to Desktop: meeting-analysis-[firstname]-YYYY-MM-DD.txt

## Known Gotchas
- Recurring meetings: filter meetings_list_transcripts by createdDateTime
- No transcript = transcription was not started during the meeting
- meetings_get_by_join_url only works if signed-in user organised the meeting

## Example Prompts
"Analyse last week's Weekly Medmin meeting for Sarah"
"What was decided in Thursday's team meeting?"
"How did James communicate in the board meeting?"
'@

Set-Content -Path (Join-Path $skillDir "SKILL.md") -Value $skillContent -Encoding UTF8
Ok "Skill file written"

$pluginsJson  = Join-Path $HOME ".claude\plugins\installed_plugins.json"
$plugins      = if (Test-Path $pluginsJson) { Get-Content $pluginsJson -Raw | ConvertFrom-Json }
                else { [PSCustomObject]@{ version = 2; plugins = [PSCustomObject]@{} } }
$key          = "$PluginName@local-plugins"
$now          = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.000Z")
$installPath  = Join-Path $HOME ".claude\plugins\cache\local-plugins\$PluginName\$PluginVersion"
$pluginRecord = @([PSCustomObject]@{ scope="user"; installPath=$installPath; version=$PluginVersion; installedAt=$now; lastUpdated=$now })

if ($plugins.plugins.PSObject.Properties.Name -contains $key) { $plugins.plugins.$key = $pluginRecord }
else { $plugins.plugins | Add-Member -NotePropertyName $key -NotePropertyValue $pluginRecord }

$plugins | ConvertTo-Json -Depth 10 | Set-Content $pluginsJson -Encoding UTF8
Ok "Teams Meeting Analyser skill installed"

# ── Sign-in wizard ─────────────────────────────────────────────────────────────
Header "Signing in to Microsoft 365"
Write-Host ""
Write-Host "  A browser window will open - sign in with your medmin.co.uk account." -ForegroundColor White
Write-Host "  If it doesn't open automatically, use the URL printed below." -ForegroundColor Gray
Write-Host ""

$env:AZURE_CLIENT_ID = $AzureClientId
$env:AZURE_TENANT_ID = $AzureTenantId
node $DistFile --setup

Write-Host ""
Ok "All done! Restart Claude Code and you're ready."
Write-Host ""
Stop-Transcript | Out-Null
Read-Host "  Press Enter to close"
