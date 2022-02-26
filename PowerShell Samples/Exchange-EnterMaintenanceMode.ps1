# Import Exchange module and internal utility functions.
. 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'
. .\CCI-UtilityFunctions.ps1

# Setup task sequence logging.
$LogFile = "C:\TaskSequencePostCheck.txt"
$LogArchive = "C:\TaskSequencePostCheck.old"

if (Test-Path $LogFile) {
    if (Test-Path $LogArchive) {
        (Get-Content -Path $LogArchive) + (Get-Content -Path $LogFile) | Set-Content -Path $LogArchive
        Clear-Content -Path $LogFile
    }
    else {
        Rename-Item -Path $LogFile -NewName "TaskSequencePostCheck.old" -Force
        New-Item $LogFile -ItemType File
    }
}
else {
    New-Item $LogFile -ItemType File
}

# Begin 'enter maintenance mode' script.
Write-CCIScriptLog -Message "-= BEGIN $($MyInvocation.MyCommand.Name) =-" -LogFile $LogFile

Connect-ExchangeServer -auto -ClientApplication:ManagementShell

$ExchangeServer = $env:computername

# Determine if this server is a member of a DAG (not a standalone Client Access Server)
[boolean]$hasDatabases = (Get-DatabaseAvailabilityGroup | Select-Object -ExpandProperty Servers) -contains $ExchangeServer

# Get random Exchange server on opposite campus for mail traffic redirection.
if ($ExchangeServer -like "BL-*") {
    if ($env:USERDNSDOMAIN -like "*testads*") {
        $IUPUIExchangeServers = Get-ExchangeServer | Where-Object { $_.name -like "in-cci-tsexch*" }
    }
    else {
        $IUPUIExchangeServers = Get-ExchangeServer | Where-Object { $_.name -like "in-cci-d*" }
    }
    $TempRedirectTargetServer = $IUPUIExchangeServers | Get-Random | Select-Object -ExpandProperty FQDN
    Write-CCIScriptLog -Message "BL server. IN server will be used for mail redirection: $TempRedirectTargetServer" -LogFile $LogFile
}
else {
    if ($env:USERDNSDOMAIN -like "*testads*") {
        $IUBExchangeServers = Get-ExchangeServer | Where-Object { $_.name -like "bl-cci-tsexch*" }
    }
    else {
        $IUBExchangeServers = Get-ExchangeServer | Where-Object { $_.name -like "bl-cci-d*" }
    }
    $TempRedirectTargetServer = $IUBExchangeServers | Get-Random | Select-Object -ExpandProperty FQDN
    Write-CCIScriptLog -Message "IN server. BL server will be used for mail redirection: $TempRedirectTargetServer" -LogFile $LogFile
}

# Drain transport roles.
Write-CCIScriptLog -Message "Draining transport roles." -LogFile $LogFile
try {
    Set-ServerComponentState $ExchangeServer -Component HubTransport -State Draining -Requester Maintenance -ErrorAction Stop
} catch {
    Write-CCIScriptLog -Message "Could not set HubTransport component to draining. Error: $($Error[0].Exception.Message)" -LogLevel ERROR -LogFile $LogFile
    Set-TSFailureFlag
    exit 1
}
Restart-Service MSExchangeTransport
Restart-Service MSExchangeFrontEndTransport

# Redirect traffic away from server.
Write-CCIScriptLog -Message "Redirecting traffic to $TempRedirectTargetServer." -LogFile $LogFile
Redirect-Message -Server $ExchangeServer -Target $TempRedirectTargetServer -confirm:$false
Write-CCIScriptLog -Message "Setting UMCallRouter to Draining." -LogFile $LogFile
try {
    Set-ServerComponentState $ExchangeServer -Component UMCallRouter -State Draining -Requester Maintenance -ErrorAction Stop
} catch {
    Write-CCIScriptLog -Message "Could not set UMCallRouter component to draining. $($Error[0].Exception.Message)" -LogLevel ERROR -LogFile $LogFile
    Set-TSFailureFlag
    exit 1
}

if ($hasDatabases) {
    # Suspend cluster service.
    Write-CCIScriptLog -Message "Suspending cluster node." -LogFile $LogFile
    Suspend-ClusterNode $ExchangeServer

    # Disable automatic DB activation.
    Write-CCIScriptLog -Message "Disabling automatic DB activation." -LogFile $LogFile
    Set-MailboxServer $ExchangeServer -DatabaseCopyAutoActivationPolicy Blocked

    # Move Active DBs to another server one at a time to minimize user impact
    $ActiveDBs = Get-MailboxDatabaseCopyStatus -Server $ExchangeServer -Active
    Write-CCIScriptLog -Message "Beginning DB moves. $($ActiveDBs.Count) active DBs at start." -LogFile $LogFile
    foreach ($DB in $ActiveDBs) {
        Write-CCIScriptLog -Message "$($DB.DatabaseName) - Move Start" -LogFile $LogFile
        $Count = 1
        while ($Count -lt 5) {
            # Find the server housing the 3rd preferred copy of the target mailbox database and mark it for move preference.
            # The 3rd preference DB copy is always opposite campus.
            $TargetMoveServer = [regex]::Match((Get-MailboxDatabase $DB.DatabaseName | Select-Object -ExpandProperty ActivationPreference | Where-Object { $_ -like "*3]" }), "[A-Z]+[^,]+") | Select-Object -ExpandProperty Value
            Write-CCIScriptLog -Message "$($DB.DatabaseName) - Moving to $TargetMoveServer (attempt $Count)" -LogFile $LogFile
            $MoveAction = Move-ActiveMailboxDatabase -Identity $DB.DatabaseName -ActivateOnServer $TargetMoveServer
            if ($MoveAction.Status -eq 'Succeeded') {
                Write-CCIScriptLog -Message "$($DB.DatabaseName) - Move Finished to $TargetMoveServer" -LogFile $LogFile
                Start-Sleep 60
                break
            }
            Start-Sleep 30
            $Count++
        }
        if ($Count -ge 5) {
            # If the move to the 3rd preferred copy fails for 2.5 minutes, another available copy is determined at random and the move is tried again.
            Write-CCIScriptLog -Message "$($DB.DatabaseName) - Move to $TargetMoveServer failed. Last error detail:" -LogLevel ERROR -LogFile $LogFile
            $MoveAction.ErrorMessage | Out-File -FilePath $LogFile -Append
            Write-CCIScriptLog -Message "$($DB.DatabaseName) - Trying to move to any available copy instead" -LogLevel WARNING -LogFile $LogFile
            $MoveAction = Move-ActiveMailboxDatabase -Identity $DB.DatabaseName
            if ($MoveAction -eq 'Succeeded') {
                Write-CCIScriptLog -Message "$($DB.DatabaseName) - Move Finished to $($MoveAction.ActiveServerAtEnd)" -LogFile $LogFile
                Start-Sleep 60
            }
            else {
                Write-CCIScriptLog -Message "$($DB.DatabaseName) - Move failed. Last error detail:" -LogLevel ERROR -LogFile $LogFile
                $MoveAction.ErrorMessage | Out-File -FilePath $LogFile -Append
            }
        }
    }

    # Disable all DB activation and force move any lingering DB copies.
    Write-CCIScriptLog -Message "Enabling DatabaseCopyActivationDisabledAndMoveNow" -LogFile $LogFile
    Set-MailboxServer $ExchangeServer -DatabaseCopyActivationDisabledAndMoveNow $True
}

# Verify that all transport queues have been drained. Ignore Shadow, Poison, and delivery to 2010 servers.
Write-CCIScriptLog -Message "Verifying message queue lengths are 0..." -LogFile $LogFile
$count = 0
while ($count -lt 5) {
    $MessageQueueLength = Get-Queue -Server $ExchangeServer | Where-Object { ($_.Identity -notlike "*\Shadow\*") -and ($_.Identity -notlike "*\Poison") -and ($_.DeliveryType -ne 'SmtpRelayToMailboxDeliveryGroup')} | Measure-Object -property MessageCount -Sum | Select-Object -ExpandProperty Sum
    if ($MessageQueueLength -ne 0) {
        Write-CCIScriptLog -Message "- Cumulative message queue length is $MessageQueueLength. Waiting 1 minute before re-evaluating." -LogFile $LogFile
        Start-Sleep -Seconds 60
    }
    else {
        Write-CCIScriptLog -Message "- Cumulative message queue length confirmed 0." -LogFile $LogFile
        break
    }
    $count++
}

if ($MessageQueueLength -ne 0) {
    while ($MessageQueueLength -ne 0) {
        Write-CCIScriptLog -Message "Message Queues are still not fully drained. Sending email notification to admins, retrying in 15 minutes." -LogLevel ERROR -LogFile $LogFile
        Get-Queue -Server $ExchangeServer | Where-Object { ($_.Identity -notlike "*\Shadow\*") -and ($_.Identity -notlike "*\Poison")} | Select-Object Identity,DeliveryType,Status,MessageCount | Out-File -FilePath $LogFile -Append
        $StuckMessages = Get-Message -Server $ExchangeServer
        $StuckMessages | Out-File -FilePath $LogFile -Append
        $StuckMessagesTable = $StuckMessages | Select-Object Identity,DateReceived,FromAddress,Status,Subject,DeferReason,Recipients | ConvertTo-Html -Fragment
        $MailProperties = @{
            SMTPServer = "mail-relay.iu.edu"
            From = "ads-admin@iu.edu"
            To = "cci-alerts@exchange.iu.edu"
            Subject = "Message Queue Stuck Draining on $ExchangeServer"
            BodyAsHTML = $true
            Body = @"
<body>
$ExchangeServer is draining its message queues to enter maintenance mode. However, some messages are not draining.<br>
Please review and either manually remove the messages or create the following file on the server to override the check: C:\Temp\Continue.txt<br>
<br>
$StuckMessagesTable
</body>
"@
        }
        Send-MailMessage @MailProperties
        Start-Sleep -Seconds 900
        if (Test-Path -Path "C:\Temp\Continue.txt") {
            Write-CCIScriptLog -Message "Manual override detected for message queue check. Setting queue length to 0 to skip check." -LogLevel WARNING -LogFile $LogFile
            Remove-Item -Path "C:\Temp\Continue.txt" -Force
            $MessageQueueLength = 0
        } else {
            $MessageQueueLength = Get-Queue -Server $ExchangeServer | Where-Object { ($_.Identity -notlike "*\Shadow\*") -and ($_.Identity -notlike "*\Poison") -and ($_.DeliveryType -ne 'SmtpRelayToMailboxDeliveryGroup')} | Measure-Object -property MessageCount -Sum | Select-Object -ExpandProperty Sum
        }
    }
    Write-CCIScriptLog -Message "Message queue length detected as 0. Continuing to enter maintenance mode." -LogFile $LogFile
}

# Set service wide maintenance mode.
Write-CCIScriptLog -Message "Putting Exchange into server wide maintenance mode" -LogFile $LogFile
try {
    Set-ServerComponentState $ExchangeServer -Component ServerWideOffline -State Inactive -Requester Maintenance -ErrorAction Stop
} catch {
    Write-CCIScriptLog -Message "Could not set ServerWideOffline to Inactive state. Error: $($Error[0].Exception.Message)" -LogLevel ERROR -LogFile $LogFile
    Set-TSFailureFlag
    exit 1
}

## Exchange Maintenance Mode - VERIFY

# Verify the server has been placed in maintenance mode.
Write-CCIScriptLog -Message "Verifying server is in maintenance mode..." -LogFile $LogFile
$ComponentStates = Get-ServerComponentState $ExchangeServer
$ServerWideOfflineState = $ComponentStates | Where-Object { $_.component -eq "ServerWideOffline" } | Select-Object -ExpandProperty state
if ($ServerWideOfflineState -eq "Active") {
    Write-TaskSequenceVariable -VariableName "OSDTSShouldContinue" -VariableValue $False
    Write-CCIScriptLog -Message "Exchange Server not detected in maintenance mode." -LogLevel ERROR -LogFile $LogFile
    Write-CCIScriptLog -Message "Error in Script: $($MyInvocation.PositionMessage)" -LogLevel ERROR -LogFile $LogFile
    $ComponentStates | Out-File -FilePath $LogFile -Append
    Set-TSFailureFlag
    exit 1
}

if ($hasDatabases) {
    Write-CCIScriptLog -Message "Verifying that there are no active DB copies..." -LogFile $LogFile
    
    # Verify the server is not hosting any active database copies
    $count = 0
    $empty = $false
    do {
        $MountedMailboxDBs = Get-MailboxDatabaseCopyStatus $ExchangeServer -Active -ErrorAction SilentlyContinue
        if ($MountedMailboxDBs.Count -ne 0) {
            Start-Sleep 300
        } else {
            $empty = $true
            break
        }
    } while ($count -le 5)

    if (!($Empty)) {
        Write-TaskSequenceVariable -VariableName "OSDTSShouldContinue" -VariableValue $False
        Write-CCIScriptLog -Message "Exchange Server still has mounted mailbox copies." -LogLevel ERROR -LogFile $LogFile
        Write-CCIScriptLog -Message "Error in Script: $($MyInvocation.PositionMessage)" -LogLevel ERROR -LogFile $LogFile
        $MountedMailboxDBs | Out-File -FilePath $LogFile -Append
        Set-TSFailureFlag
        exit 1
    }

    # Verify the node is paused
    $ClusterNodeStatus = Get-ClusterNode $ExchangeServer | Select-Object -ExpandProperty State
    if ($ClusterNodeStatus -eq "Up") {
        Write-TaskSequenceVariable -VariableName "OSDTSShouldContinue" -VariableValue $False
        Write-CCIScriptLog -Message "Cluster Node status is still Up on server." -LogLevel ERROR -LogFile $LogFile
        Write-CCIScriptLog -Message "Error in Script: $($MyInvocation.PositionMessage)" -LogLevel ERROR -LogFile $LogFile
        $ClusterNodeStatus | Out-File -FilePath $LogFile -Append
        Set-TSFailureFlag
        exit 1
    }
}

Write-CCIScriptLog -Message "-= END $($MyInvocation.MyCommand.Name) =-" -LogFile $LogFile