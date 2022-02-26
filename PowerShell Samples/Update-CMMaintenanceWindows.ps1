<#
    .SYNOPSIS
        Updates the monthly Patch Wednesday maintenance windows

    .DESCRIPTION
        The second Wednesday of every month does not necessarily follow the second Tuesday. This script and associated
        task will ensure that it stays scheduled for the Wednesday following Patch Tuesday.
    
    .NOTES        
        Script will be automatically ran the first of every month at 7:00am.
#>

function Get-CCIPatchTuesday {
    param(
        [int]$Month,
        [int]$Year
    )
    [datetime]$StartDate = $Month.ToString() + "/1/" + $Year.ToString()
    $8thOfMonth = $StartDate.AddDays(7)
    if ($8thOfMonth.DayOfWeek -eq "Tuesday") {
        return $8thOfMonth
    }
    else {
        $count = 1
        while ($8thOfMonth.AddDays($count).DayOfWeek -ne "Tuesday") {
            $count++
        }
        return $8thOfMonth.AddDays($count)
    }
}

Import-Module $env:SMS_ADMIN_UI_PATH.Replace("\bin\i386","\bin\configurationmanager.psd1")
Set-Location CCI:

# Determine date for "Patch Wednesday" and get all "Update Group*" device collections
$PatchWednesdayDate = (Get-CCIPatchTuesday -Month ([datetime]::Now).Month -Year ([datetime]::Now).Year).AddDays(1)
$UpdateGroups = Get-CMCollection | Where-Object { $_.Name -like "Update Group*" }
[string]$MWDetails = ""

foreach ($UpdateGroup in $UpdateGroups) {
    if ($UpdateGroup.Name -like "*Wed,*") {
        # This depends on the update group collections being named consistently; e.g. "Update Group - 2nd Wed, XPM-YAM"
        $RawRange = $UpdateGroup.Name.Split(",")[1].Trim()
        $RawStartTime = [DateTime]$RawRange.Split("-")[0]
        $RawEndTime = [DateTime]$RawRange.Split("-")[1]
        # If the start time per the group name is midnight, this will adjust the actual window start time to 11:59PM to make things a bit easier
        if ($RawStartTime.Hour -eq 0) {
            $RawStartTime = ($RawStartTime.AddDays(1)).AddMinutes(-1)
        }
        # The logic here is that any Wednesday night maintenance windows will either be PM->PM or PM->AM, so <12 means morning/AM next day.
        if ($RawEndTime.Hour -lt 12) {
            $RawEndTime = $RawEndTime.AddDays(1)
        }

        # Determine intended duration of maintenance window, then use that information to create Patch Wednesday start/end times
        $Duration = $RawEndTime - $RawStartTime
        $StartTime = ($PatchWednesdayDate.AddHours($RawStartTime.Hour)).AddMinutes($RawStartTime.Minute)
        $EndTime = $StartTime.AddHours($Duration.Hours)
        
        # Create new CM schedule object, remove old maintenance window, and create a new one for the current month's Patch Wednesday
        $MWSchedule = New-CMSchedule -Start $StartTime -End $EndTime -Nonrecurring
        $ExistingMaintenanceWindow = Get-CMMaintenanceWindow -CollectionID $UpdateGroup.CollectionID
        if ($ExistingMaintenanceWindow) {
            $ExistingMaintenanceWindow | ForEach-Object {
                Remove-CMMaintenanceWindow -CollectionID $UpdateGroup.CollectionID -MaintenanceWindowName $_.Name -Force
            }
        }
        New-CMMaintenanceWindow -Schedule $MWSchedule -Name "Patch Wednesday, $(Get-Date $StartTime -Format hhtt) - $(Get-Date $EndTime -Format hhtt)" -CollectionId $UpdateGroup.CollectionID -ApplyTo SoftwareUpdatesOnly
        Start-Sleep -Seconds 4
        $UpdatedMaintenanceWindow = Get-CMMaintenanceWindow -CollectionID $UpdateGroup.CollectionID
        $MWDetails += "<tr><td>$($UpdateGroup.Name)</td><td>$($UpdatedMaintenanceWindow.Description)</td><td>$($UpdatedMaintenanceWindow.Duration / 60) hours</td></tr>"
    }
}

# Send an email with updated maintenance window information so script's work can be easily validated
$EmailTo = "ccisccm@iu.edu"
$EmailFrom = "sccmalert@iu.edu"
$EmailSubject = "SCCM 2nd Wednesday Maintenance Windows Updated"
$BodyStart = @"
<body>The SCCM 2nd Wednesday maintenance windows for production systems have been updated. Please review list for accuracy:<br>
<br>
<table border=1>
<tr><td><h3>Collection Name</h3></td><td><h3>Maintenance Window</h3></td><td><h3>Duration</h3></td></tr>
"@
$BodyEnd = $MWDetails + "</table></body>"
$EmailBody = $BodyStart + $BodyEnd
Send-MailMessage -BodyAsHtml -Body $EmailBody -From $EmailFrom -To $EmailTo -Subject $EmailSubject -SmtpServer "mail-relay.iu.edu" -UseSsl