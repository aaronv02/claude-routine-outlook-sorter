#Requires -Version 5.1
<#
.SYNOPSIS
  Registers the two routines as Windows scheduled tasks on this machine.

.DESCRIPTION
  Everything then lives on one computer: no cloud runner, no secret stored anywhere
  but this machine, nothing that depends on another person's account.

  Two tasks are created:

    Outlook Sorter - Sort       hourly, 8am to 6pm, weekdays
    Outlook Sorter - Weekly     Friday 3pm

  Both run only while this user is logged on. That is deliberate rather than a
  limitation worked around: running as a service would mean storing the account
  password, and the weekly summary already reports the week it missed if a run does
  not happen, so a laptop that was closed on Friday afternoon is a case the tool
  handles rather than a case it needs to prevent.

  Requires the Claude Code CLI, which is what actually performs the classification
  and writes the summary. If it is missing, this offers to install it.

.PARAMETER StartHour
  First hourly sorting run. Default 8.

.PARAMETER EndHour
  Last hourly sorting run. Default 18.

.PARAMETER SummaryTime
  When the weekly summary runs on Friday, 24-hour HH:mm. Default 15:00.

.EXAMPLE
  .\scripts\install-tasks.ps1

.EXAMPLE
  .\scripts\install-tasks.ps1 -StartHour 7 -EndHour 19 -SummaryTime 16:30
#>
[CmdletBinding()]
param(
    [int]$StartHour = 8,
    [int]$EndHour = 18,
    [string]$SummaryTime = '15:00'
)

$ErrorActionPreference = 'Stop'

$SortTaskName = 'Outlook Sorter - Sort'
$WeeklyTaskName = 'Outlook Sorter - Weekly'

function Write-Step { param([string]$T) Write-Host "`n$T" -ForegroundColor Cyan }
function Write-Ok   { param([string]$T) Write-Host "  OK  $T" -ForegroundColor Green }
function Write-Warn { param([string]$T) Write-Host "  !   $T" -ForegroundColor Yellow }

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Write-Host "`nScheduling the Outlook routines" -ForegroundColor White
Write-Host "Project: $ProjectRoot"

# ---------------------------------------------------------------------------
# Preconditions
# ---------------------------------------------------------------------------

Write-Step 'Checking prerequisites...'

if (-not (Test-Path (Join-Path $ProjectRoot '.env'))) {
    Write-Warn 'No .env found - this project has not been set up yet.'
    Write-Host "`n  Run this first, in $ProjectRoot :`n"
    Write-Host "      npm run setup`n"
    exit 1
}
Write-Ok 'Project is set up'

$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) { Write-Warn 'Node.js is not on PATH.'; exit 1 }
Write-Ok "Node at $($node.Source)"

if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
    Write-Warn 'The Claude Code CLI is not installed.'
    Write-Host @'

  It is what actually reads the mail and decides the labels. Without it the
  scheduled tasks would run and do nothing.

'@
    $answer = Read-Host '  Install it now with npm? [Y/n]'
    if ($answer -match '^[Nn]') {
        Write-Host "`n  Install it yourself with:`n"
        Write-Host "      npm install -g @anthropic-ai/claude-code`n"
        exit 1
    }

    & npm install -g '@anthropic-ai/claude-code'
    if ($LASTEXITCODE -ne 0) {
        Write-Warn 'That failed. On some machines global npm installs need an administrator.'
        exit 1
    }

    if (-not (Get-Command claude -ErrorAction SilentlyContinue)) {
        Write-Warn 'Installed, but `claude` is still not on PATH.'
        Write-Host '      Close this window, open a new PowerShell, and run this script again.'
        exit 1
    }
}
Write-Ok 'Claude Code CLI present'

Write-Host @'

  One thing to confirm by hand: the CLI needs to be signed in to Claude, once.
  If you have not already, run `claude` in this folder and complete the sign-in
  before relying on the schedule.

'@

# ---------------------------------------------------------------------------
# Register
# ---------------------------------------------------------------------------

Write-Step 'Registering scheduled tasks...'

# npm.cmd resolved absolutely: Task Scheduler does not run through a shell and does
# not inherit an interactive PATH, so a bare "npm" is not found at run time - the
# task then fails with an opaque 0x1 that looks like the script's fault.
$npmCmd = (Get-Command npm.cmd -ErrorAction SilentlyContinue)
if (-not $npmCmd) { $npmCmd = Get-Command npm -ErrorAction Stop }
$npmPath = $npmCmd.Source

function Register-RoutineTask {
    param(
        [string]$Name,
        [string]$Script,
        [Microsoft.Management.Infrastructure.CimInstance[]]$Triggers,
        [string]$Description
    )

    $action = New-ScheduledTaskAction -Execute $npmPath `
        -Argument "run --silent $Script" `
        -WorkingDirectory $ProjectRoot

    # Interactive so it runs as this logged-on user with their own Claude sign-in.
    # A service account would need a stored password and would not have it.
    $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
        -LogonType Interactive -RunLevel Limited

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

    # StartWhenAvailable catches up a run missed because the machine was asleep.
    # IgnoreNew means a long run never has a second copy started on top of it,
    # which would have two processes writing the same mailbox state.

    if (Get-ScheduledTask -TaskName $Name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $Name -Confirm:$false
    }

    Register-ScheduledTask -TaskName $Name `
        -Action $action -Trigger $Triggers -Principal $principal -Settings $settings `
        -Description $Description | Out-Null

    Write-Ok $Name
}

$hours = $EndHour - $StartHour
if ($hours -lt 1) { Write-Warn 'EndHour must be later than StartHour.'; exit 1 }

$sortTrigger = New-ScheduledTaskTrigger -Daily -At ("{0:00}:00" -f $StartHour)
$sortTrigger.Repetition = (New-ScheduledTaskTrigger -Once -At ("{0:00}:00" -f $StartHour) `
    -RepetitionInterval (New-TimeSpan -Hours 1) `
    -RepetitionDuration (New-TimeSpan -Hours $hours)).Repetition

Register-RoutineTask -Name $SortTaskName -Script 'sort' -Triggers @($sortTrigger) `
    -Description 'Sorts new Outlook mail into categories and learns from corrections.'

$weeklyTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Friday -At $SummaryTime
Register-RoutineTask -Name $WeeklyTaskName -Script 'summary' -Triggers @($weeklyTrigger) `
    -Description 'Writes the end-of-week summary of what is still waiting on a reply.'

# ---------------------------------------------------------------------------

Write-Host @"

$(' ' * 0)Done.

  Sorting          hourly, ${StartHour}:00 to ${EndHour}:00, every day
  Weekly summary   Fridays at $SummaryTime

  Both run only while you are logged on. A missed run is picked up when the
  machine next wakes, and the weekly summary reports the week it missed rather
  than a week that has barely started.

  Logs:      $ProjectRoot\routine\.local\logs\
  Try one:   npm run sort
  Remove:    .\scripts\uninstall-tasks.ps1
  Inspect:   open Task Scheduler and look for "Outlook Sorter"

"@ -ForegroundColor White
