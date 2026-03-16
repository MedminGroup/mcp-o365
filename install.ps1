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
# Finds Claude automatically, registers .dxt file type, installs the extension.
# ──────────────────────────────────────────────────────────────────────────────
param([string]$Source = "")

# ── Logging ────────────────────────────────────────────────────────────────────
$LogFile = Join-Path $env:TEMP "medmin-install-log.txt"
Start-Transcript -Path $LogFile -Force | Out-Null
Write-Host "  [log] Full log: $LogFile" -ForegroundColor DarkGray

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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
    try {
        $releases = Invoke-RestMethod "https://nodejs.org/dist/index.json" `
            -Headers @{"User-Agent"="mcp-o365-installer"}
        $lts     = ($releases | Where-Object { $_.lts -ne $false } | Select-Object -First 1)
        $version = $lts.version
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

    $userPath = [Environment]::GetEnvironmentVariable("PATH", "User")
    if ($null -eq $userPath -or $userPath -notlike "*$NodeDir*") {
        [Environment]::SetEnvironmentVariable("PATH", "$NodeDir;$userPath", "User")
    }
    $env:PATH = "$NodeDir;$env:PATH"
    Ok "Node.js $version installed to $NodeDir"
}

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

$nodeExe = Get-Command node -ErrorAction SilentlyContinue
if (-not $nodeExe) { throw "node.exe not found on PATH after install. PATH is: $env:PATH" }
Ok "node is at $($nodeExe.Source)"

# ── Download MCP server (needed for sign-in wizard) ────────────────────────────
Header "Downloading MCP server"

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
        $release = Invoke-RestMethod -Uri $ReleasesUrl -Headers @{"User-Agent"="mcp-o365-installer"}
    } catch {
        Fail "Could not reach GitHub API: $_"
    }
    $asset = $release.assets | Where-Object { $_.name -eq "index.js" } | Select-Object -First 1
    if (-not $asset) { Fail "index.js not found in latest release assets." }
    Info "Downloading $($asset.browser_download_url)"
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $DistFile -UseBasicParsing
}
Ok "MCP server ready at $DistFile"

# ── Find Claude desktop app ────────────────────────────────────────────────────
Header "Finding Claude desktop app"

$ClaudeExe = $null

# 1. Microsoft Store install — use Get-AppxPackage (no admin needed, works on all Store installs)
$pkg = Get-AppxPackage -Name "*Claude*" -ErrorAction SilentlyContinue
if ($pkg) {
    $candidate = Join-Path $pkg.InstallLocation "app\Claude.exe"
    if (Test-Path $candidate) {
        $ClaudeExe = $candidate
        Ok "Found Claude (Store): $ClaudeExe"
    }
}

# 2. Standard direct-install paths
if (-not $ClaudeExe) {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA "AnthropicClaude\Claude.exe"),
        (Join-Path $env:PROGRAMFILES "Claude\Claude.exe")
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) { $ClaudeExe = $c; Ok "Found Claude: $ClaudeExe"; break }
    }
}

if (-not $ClaudeExe) {
    Warn "Claude desktop app not found. Please install it from https://claude.ai/download"
    Warn "After installing Claude, re-run this installer."
    Fail "Claude not found — cannot continue."
}

# ── Enable Windows Developer Mode (required for Claude extensions) ─────────────
Header "Checking Windows Developer Mode"

$devModeKey  = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock"
$devModeVal  = Get-ItemProperty -Path $devModeKey -Name "AllowDevelopmentWithoutDevLicense" -ErrorAction SilentlyContinue
$devModeOn   = ($devModeVal -and $devModeVal.AllowDevelopmentWithoutDevLicense -eq 1)

if ($devModeOn) {
    Ok "Windows Developer Mode is already enabled"
} else {
    Info "Windows Developer Mode is required for Claude extensions."
    Info "A UAC prompt will appear — click Yes to allow the change..."

    $regCmd = @"
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' -Name 'AllowDevelopmentWithoutDevLicense' -Value 1 -Type DWord -Force
Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock' -Name 'AllowAllTrustedApps' -Value 1 -Type DWord -Force
"@
    $tmpReg = Join-Path $env:TEMP "enable-devmode.ps1"
    Set-Content $tmpReg -Value $regCmd -Encoding UTF8
    Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$tmpReg`"" -Wait
    Remove-Item $tmpReg -Force -ErrorAction SilentlyContinue

    # Verify
    $devModeVal = Get-ItemProperty -Path $devModeKey -Name "AllowDevelopmentWithoutDevLicense" -ErrorAction SilentlyContinue
    if ($devModeVal -and $devModeVal.AllowDevelopmentWithoutDevLicense -eq 1) {
        Ok "Windows Developer Mode enabled"
    } else {
        Warn "Could not enable automatically. Please enable it manually:"
        Warn "  Settings -> System -> For developers -> Developer Mode -> On"
        Warn "Then re-run this installer."
        Fail "Developer Mode not enabled"
    }
}

# ── Register .dxt file association ─────────────────────────────────────────────
Header "Registering .dxt file type"

# Write to HKCU — no admin required
$null = New-Item -Path "HKCU:\Software\Classes\.dxt" -Force
Set-ItemProperty -Path "HKCU:\Software\Classes\.dxt" -Name "(default)" -Value "ClaudeDXTFile" -Force

$null = New-Item -Path "HKCU:\Software\Classes\ClaudeDXTFile" -Force
Set-ItemProperty -Path "HKCU:\Software\Classes\ClaudeDXTFile" -Name "(default)" -Value "Claude Desktop Extension" -Force

$null = New-Item -Path "HKCU:\Software\Classes\ClaudeDXTFile\DefaultIcon" -Force
Set-ItemProperty -Path "HKCU:\Software\Classes\ClaudeDXTFile\DefaultIcon" -Name "(default)" -Value "`"$ClaudeExe`",0" -Force

$null = New-Item -Path "HKCU:\Software\Classes\ClaudeDXTFile\shell\open\command" -Force
Set-ItemProperty -Path "HKCU:\Software\Classes\ClaudeDXTFile\shell\open\command" -Name "(default)" -Value "`"$ClaudeExe`" `"%1`"" -Force

Ok ".dxt files now open with Claude (double-click will work from now on)"

# ── Download and install the DXT extension ─────────────────────────────────────
Header "Installing Medmin extension into Claude"

$DxtFile = Join-Path $env:TEMP "medmin-m365.dxt"
$ReleasesUrl = "https://api.github.com/repos/$GitHubRepo/releases/latest"
try {
    $release = Invoke-RestMethod -Uri $ReleasesUrl -Headers @{"User-Agent"="mcp-o365-installer"}
} catch {
    Fail "Could not reach GitHub API: $_"
}
$dxtAsset = $release.assets | Where-Object { $_.name -eq "medmin-m365.dxt" } | Select-Object -First 1
if (-not $dxtAsset) { Fail "medmin-m365.dxt not found in latest release." }

Info "Downloading medmin-m365.dxt..."
Invoke-WebRequest -Uri $dxtAsset.browser_download_url -OutFile $DxtFile -UseBasicParsing
Ok "Downloaded extension"

Info "Opening in Claude — an install prompt will appear, click Install..."
Start-Process $ClaudeExe -ArgumentList "`"$DxtFile`""
Start-Sleep -Seconds 4
Ok "Extension sent to Claude"

# ── Install CLAUDE.md ──────────────────────────────────────────────────────────
Header "Installing Claude instructions"

$claudeMd = @'
# Medmin Microsoft 365 — Claude Instructions

You have the **mcp-o365** MCP server connected. It gives you live access to this
user's Microsoft 365 account via the Microsoft Graph API.

## Critical rules — read before every Microsoft 365 task

1. **Always use mcp-o365 tools.** Never use any built-in Microsoft plugin or
   integration. The built-in Microsoft plugin does not work on Medmin accounts.
   mcp-o365 is the only authorised integration.

2. **You have a valid token.** The user is already signed in. Do not ask them to
   sign in unless accounts_list returns an empty list.

3. **For any Teams meeting or transcript request**, always follow this sequence:
   - accounts_list -> confirm the account
   - calendar_list_events -> find the meeting and its joinWebUrl
   - meetings_get_by_join_url -> get the meeting ID
   - meetings_list_transcripts -> list available transcripts
   - meetings_get_transcript -> download the VTT content
   Then analyse the transcript.

4. **Never guess or fabricate transcript content.** Always fetch it with the tools.

5. **If a tool call fails**, report the exact error. Do not fall back to the
   built-in Microsoft plugin.
'@

Set-Content -Path (Join-Path $HOME "CLAUDE.md") -Value $claudeMd -Encoding UTF8
Ok "CLAUDE.md installed"

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
Ok "All done!"
Write-Host "  If Claude opened an install prompt, click Install." -ForegroundColor White
Write-Host "  Then restart Claude and ask: `"what can you do?`"" -ForegroundColor White
Write-Host ""
Stop-Transcript | Out-Null
Read-Host "  Press Enter to close"
