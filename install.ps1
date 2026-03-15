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
# Configures the Claude desktop app — no Claude Code CLI needed.
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

# ── Configure Claude desktop app ───────────────────────────────────────────────
Header "Configuring Claude desktop app"

# Claude installed from Microsoft Store uses a sandboxed path under LocalAppData\Packages.
# Claude installed directly uses the standard %APPDATA%\Claude path.
# We probe both and use whichever exists, preferring the Store path.
$ClaudeConfigDir = $null

# 1. Try known Store package name (publisher hash is consistent for Anthropic's app)
$storePath = Join-Path $env:LOCALAPPDATA "Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude"
if (Test-Path $storePath) {
    $ClaudeConfigDir = $storePath
    Info "Found Microsoft Store install at $storePath"
}

# 2. Wildcard search in case publisher hash differs
if (-not $ClaudeConfigDir) {
    $pkgDir = Join-Path $env:LOCALAPPDATA "Packages"
    $match  = Get-ChildItem $pkgDir -Filter "Claude_*" -Directory -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($match) {
        $candidate = Join-Path $match.FullName "LocalCache\Roaming\Claude"
        if (Test-Path $candidate) {
            $ClaudeConfigDir = $candidate
            Info "Found Microsoft Store install at $candidate"
        }
    }
}

# 3. Fall back to standard direct-install path
if (-not $ClaudeConfigDir) {
    $ClaudeConfigDir = Join-Path $env:APPDATA "Claude"
    Info "Using standard install path at $ClaudeConfigDir"
}

$ClaudeConfig = Join-Path $ClaudeConfigDir "claude_desktop_config.json"
New-Item -ItemType Directory -Force -Path $ClaudeConfigDir | Out-Null

# Create a minimal config if it doesn't exist
if (-not (Test-Path $ClaudeConfig)) {
    Info "Creating new config at $ClaudeConfig"
    '{"mcpServers":{}}' | Set-Content $ClaudeConfig -Encoding UTF8
}

$config = Get-Content $ClaudeConfig -Raw | ConvertFrom-Json
$existed = $false

if (-not $config.PSObject.Properties["mcpServers"]) {
    $config | Add-Member -NotePropertyName "mcpServers" -NotePropertyValue ([PSCustomObject]@{})
}
if ($config.mcpServers.PSObject.Properties["mcp-o365"]) { $existed = $true }

# Resolve full path to node.exe so Claude desktop app doesn't need node on its PATH
$NodeExePath = (Get-Command node).Source

$entry = [PSCustomObject]@{
    command = $NodeExePath
    args    = @($DistFile)
    env     = [PSCustomObject]@{ AZURE_CLIENT_ID = $AzureClientId; AZURE_TENANT_ID = $AzureTenantId }
}
if ($existed) { $config.mcpServers."mcp-o365" = $entry }
else          { $config.mcpServers | Add-Member -NotePropertyName "mcp-o365" -NotePropertyValue $entry }

$config | ConvertTo-Json -Depth 10 | Set-Content $ClaudeConfig -Encoding UTF8
Ok ($(if ($existed) { "Updated" } else { "Registered" }) + " mcp-o365 in Claude desktop app")

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
Ok "All done! Restart the Claude app and you're ready."
Write-Host "  Open Claude and ask: `"what can you do?`"" -ForegroundColor White
Write-Host ""
Stop-Transcript | Out-Null
Read-Host "  Press Enter to close"
