#Requires -Version 5.1
<#
.SYNOPSIS
  Downloads and installs the Outlook sorter, then hands off to the setup wizard.

.DESCRIPTION
  One paste, so there is nothing to explain to whoever is sitting at the machine.
  It checks Node, downloads the project, installs its dependencies, and starts
  `npm run setup`.

  Three things it deliberately does NOT do:

    - Install Node. That needs an administrator on a managed laptop, and silently
      installing a runtime on someone's work machine is not this script's business.
      It tells you where to get it and stops.
    - Sign in. The mailbox owner does that herself in the wizard - that consent is
      the entire security model.
    - Restart Claude Desktop. The wizard says when, and a human has to quit it from
      the system tray.

.PARAMETER Path
  Where to install. Defaults to C:\outlook-sorter.

  Deliberately not under Documents or Desktop: on Microsoft 365 machines those are
  usually redirected into OneDrive, and node_modules there means thousands of files
  churning through sync, plus file-lock failures during install. It also must be
  somewhere permanent - Claude Desktop stores absolute paths into this folder, so
  moving or deleting it later silently breaks the connection.

.PARAMETER SkipSetup
  Install only; don't start the wizard.

.EXAMPLE
  irm https://raw.githubusercontent.com/aaronv02/claude-routine-outlook-sorter/main/install.ps1 | iex

.EXAMPLE
  .\install.ps1 -Path D:\tools\outlook-sorter
#>
[CmdletBinding()]
param(
    [string]$Path = 'C:\outlook-sorter',
    [switch]$SkipSetup
)

$ErrorActionPreference = 'Stop'

# Stock Windows PowerShell 5.1 still negotiates TLS 1.0 by default, which GitHub
# refuses. Without this the download fails with an unhelpful "could not create
# SSL/TLS secure channel".
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Invoke-WebRequest's progress bar makes large downloads dramatically slower in
# 5.1 - a known bug, and this is a several-megabyte zip.
$ProgressPreference = 'SilentlyContinue'

$RepoZip = 'https://github.com/aaronv02/claude-routine-outlook-sorter/archive/refs/heads/main.zip'
$InnerFolder = 'claude-routine-outlook-sorter-main'

function Write-Step { param([string]$Text) Write-Host "`n$Text" -ForegroundColor Cyan }
function Write-Ok   { param([string]$Text) Write-Host "  OK  $Text" -ForegroundColor Green }
function Write-Warn { param([string]$Text) Write-Host "  !   $Text" -ForegroundColor Yellow }

Write-Host "`nOutlook Sorter - installer" -ForegroundColor White

# ---------------------------------------------------------------------------
# Node
# ---------------------------------------------------------------------------

Write-Step 'Checking for Node.js...'

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
    Write-Warn 'Node.js is not installed.'
    Write-Host @'

  Install the LTS version from:

      https://nodejs.org

  On a work laptop this may ask for an administrator password. If you do not have
  one, whoever manages the computer has to install it. Nothing here works without
  it.

  Once Node is installed, close this window, open a new PowerShell, and run this
  installer again.
'@
    exit 1
}

$version = (& node --version) -replace '^v', ''
$major = [int]($version -split '\.')[0]
if ($major -lt 20) {
    Write-Warn "Node $version is too old - version 20 or newer is required."
    Write-Host "`n  Install the current LTS from https://nodejs.org, then run this again.`n"
    exit 1
}
Write-Ok "Node $version"

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Warn 'npm was not found, which is unusual - it ships with Node.'
    Write-Host "`n  Reinstalling Node from https://nodejs.org should fix it.`n"
    exit 1
}

# ---------------------------------------------------------------------------
# Where it goes
# ---------------------------------------------------------------------------

Write-Step "Installing to $Path"

if ($Path -match 'OneDrive') {
    Write-Warn 'That path looks like it is inside OneDrive.'
    Write-Host '      Sync will fight with node_modules. Consider -Path C:\outlook-sorter instead.'
    $answer = Read-Host '      Continue anyway? [y/N]'
    if ($answer -notmatch '^[Yy]') { exit 1 }
}

$alreadyThere = (Test-Path $Path) -and (Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue)
if ($alreadyThere) {
    Write-Warn "$Path already exists and is not empty."
    Write-Host '      Continuing will overwrite the program files.'
    Write-Host '      Your sign-in (.env) and local files are preserved.'
    $answer = Read-Host '      Continue? [y/N]'
    if ($answer -notmatch '^[Yy]') { exit 1 }
}

# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------

Write-Step 'Downloading...'

$temp = Join-Path $env:TEMP ("outlook-sorter-" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $temp -Force | Out-Null
$zip = Join-Path $temp 'source.zip'

try {
    Invoke-WebRequest -Uri $RepoZip -OutFile $zip -UseBasicParsing
    Write-Ok ('Downloaded {0:N1} MB' -f ((Get-Item $zip).Length / 1MB))

    Expand-Archive -Path $zip -DestinationPath $temp -Force
    $extracted = Join-Path $temp $InnerFolder
    if (-not (Test-Path $extracted)) {
        throw "The archive did not contain the expected folder '$InnerFolder'."
    }

    New-Item -ItemType Directory -Path $Path -Force | Out-Null

    # Copied rather than moved, and per-item, so an existing .env and
    # routine\.local survive an upgrade. Losing the .env would mean the mailbox
    # owner has to sign in again for no reason.
    Copy-Item -Path (Join-Path $extracted '*') -Destination $Path -Recurse -Force
    Write-Ok "Files in place at $Path"
}
finally {
    Remove-Item -Path $temp -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# Dependencies
# ---------------------------------------------------------------------------

Write-Step 'Installing dependencies (this takes a minute)...'

Push-Location $Path
try {
    & npm install --no-fund --no-audit
    if ($LASTEXITCODE -ne 0) {
        Write-Warn 'npm install failed.'
        Write-Host "`n  Common causes: no internet, a proxy, or antivirus blocking npm."
        Write-Host "  You can retry by hand:`n"
        Write-Host "      cd $Path"
        Write-Host "      npm install`n"
        exit 1
    }
    Write-Ok 'Dependencies installed'

    if ($SkipSetup) {
        Write-Host "`nDone. Run the wizard when ready:`n"
        Write-Host "    cd $Path"
        Write-Host "    npm run setup`n"
        exit 0
    }

    # ---------------------------------------------------------------------
    # Hand off
    # ---------------------------------------------------------------------

    Write-Host @"

Starting setup.

The next part needs the mailbox owner: she signs in partway through, and whoever
signs in is whose mail gets sorted.

"@ -ForegroundColor White

    & npm run setup
}
finally {
    Pop-Location
}
