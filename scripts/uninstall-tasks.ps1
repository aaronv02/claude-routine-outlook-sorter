#Requires -Version 5.1
<#
.SYNOPSIS
  Removes the two scheduled tasks. Leaves the project, the sign-in, and the
  mailbox untouched.

.DESCRIPTION
  Stopping the schedule does not undo anything it did. Labels already applied stay,
  and any sender promoted into a native Outlook rule keeps being labelled on
  arrival by Outlook itself - that is the whole point of promotion, and it survives
  this script deliberately.

  To undo those too: delete the categories from Outlook's category list, and delete
  the "Inbox Steward:" rules under Outlook's own Rules settings.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$names = @('Outlook Sorter - Sort', 'Outlook Sorter - Weekly')
$removed = 0

foreach ($name in $names) {
    if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false
        Write-Host "  Removed: $name" -ForegroundColor Green
        $removed++
    }
    else {
        Write-Host "  Not found: $name" -ForegroundColor DarkGray
    }
}

Write-Host @"

$removed task(s) removed. Nothing else was changed.

Labels already applied stay where they are, and senders already promoted into
native Outlook rules keep being labelled by Outlook on arrival - that runs in the
mailbox, not here.

To undo those as well:
  - Delete the categories from Outlook's category list.
  - Delete the "Inbox Steward:" rules in Outlook's Rules settings.

To stop Claude Desktop from seeing the mailbox:  npm run disconnect

"@
