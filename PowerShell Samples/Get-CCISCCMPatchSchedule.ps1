#====================================================================================================================================================
##############################
## Get-CCISCCMPatchSchedule ##
##############################

#region Get-CCISCCMPatchSchedule

function Get-CCISCCMPatchSchedule {
    <#
        .SYNOPSIS
            Retrieves the current scheduled patching dates/time from SCCM

        .DESCRIPTION
            Retrieves the current scheduled patching dates/time from a given SCCM management point/site. 

        .PARAMETER ManagementPointFQDN
            The FQDN of the management point used for connecting to SCCM

        .PARAMETER SiteCode
            The 3 character site code of the SCCM site being queried

        .PARAMETER ShowUnassignedActiveMachines
            Instead of returning the patch schedule, the function will return any active machines that have yet to be assigned to one of the
            maintenance window collections.

        .PARAMETER IncludeSundays
            Also returns any Sunday patch collections. These are left out by default since every active machine should automatically be added
            to a Sunday patch collection.

        .EXAMPLE
            Get-CCISCCMPatchSchedule -ManagementPointFQDN bl-cci-cmmp01.ads.iu.edu -SiteCode CCI
            This will return a table of active machines in SCCM and what patch collections they belong to.
    #>

    [CmdletBinding(DefaultParameterSetName='GetPatchSchedule')]
    param(
        [Parameter(ParameterSetName="GetPatchSchedule")]
        [Parameter(ParameterSetName="GetUnassignedMachines")]
        [Parameter(ParameterSetName="IncludeSundays")]
        [Parameter(Mandatory=$True)][string]$ManagementPointFQDN,

        [Parameter(ParameterSetName="GetPatchSchedule")]
        [Parameter(ParameterSetName="GetUnassignedMachines")]
        [Parameter(ParameterSetName="IncludeSundays")]
        [Parameter(Mandatory=$True)][string]$SiteCode,

        [Parameter(ParameterSetName="GetUnassignedMachines")]
        [switch]$ShowUnassignedActiveMachines,

        [Parameter(ParameterSetName="IncludeSundays")]
        [switch]$IncludeSundays
    )

    begin {
        $ModulePath = try {
            $env:SMS_ADMIN_UI_PATH.Replace('\bin\i386','\bin\configurationmanager.psd1')
        } catch {
            "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\configurationmanager.psd1"
        }
        if ((Get-Module -Name ConfigurationManager) -as [bool]) {
            # Module is already loaded
        } else {
            if (Test-Path -Path $ModulePath) {
                Import-Module $ModulePath
            } else {
                throw "SCCM Admin Console is not installed on this machine."
            }
        }
        if (!(Test-Path -Path ($SiteCode + ":\"))) {
            try {
                New-PSDrive -Name $SiteCode -Root $ManagementPointFQDN -PSProvider AdminUI.PS.Provider\CMSite -ErrorAction Stop
            } catch {
                throw "Management Point FQDN or Site Code is incorrect."
            }
        }
        $CurrentLocation = Get-Location
        Set-Location -Path ($SiteCode + ":")
    }

    process {
        # Get non-Sunday maintenance window collections
        $MWCollections = Get-CMCollection | Where-Object { $_.name -match "Update Group -" } | Select-Object -ExpandProperty name
        if (!($IncludeSundays)) {
            $MWCollections = $MWCollections | Where-Object { $_ -notmatch "Sunday" }
        }

        # Get all active SCCM managed devices that aren't SCCM servers
        $CMSiteSystems = (Get-CMSiteSystemServer | Select-Object -ExpandProperty networkospath)
        # Regex matches the FQDN in a network path and returns the hostname, e.g. "\\iu-test-machine.ads.iu.edu\blah" => "iu-test-machine"
        $CMSiteSystems = $CMSiteSystems | ForEach-Object { [regex]::Match($_,"^\\\\(.+?)\.").Captures.Groups[1].Value }
        $ActiveDevices = Get-CMDevice | Where-Object { $null -ne $_.IsActive -and $CMSiteSystems -notcontains $_.Name } | Sort-Object name

        # Create list of devices added to maintenance window collections
        $MachinesWithMW = @()
        foreach ($MWCollection in $MWCollections) {
            $CollectionMembers = Get-CMCollectionMember -CollectionName $MWCollection | Sort-Object Name
            # Retrieve/calculate next patch days/times (some calculation is needed since there isn't a "next run time/time frame" attribute)
            switch -Wildcard ($MWCollection) {
                "*Exchange*" {
                    $PatchTuesdayCurrentMonth = Get-CCIPatchTuesday -Month (Get-Date).Month -Year (Get-Date).Year
                    if ((Get-Date) -le $PatchTuesdayCurrentMonth) {
                        $PatchStartDate = $PatchTuesdayCurrentMonth.ToString('MM-dd-yyyy')
                    } else {
                        $NextMonthDate = (Get-Date).AddMonths(1)
                        $PatchStartDate = (Get-CCIPatchTuesday -Month $NextMonthDate.Month -Year $NextMonthDate.Year).ToString('MM-dd-yyyy')
                    }
                    # Patching for Exchange servers is handled via scheduled Task Sequences, so we look to a deployment time instead of a maintenance window.
                    # The patching task sequence is the only thing deployed to the Exchange Update Group device collection.
                    $PatchStartTime = (Get-CMDeployment -CollectionName $MWCollection).DeploymentTime.ToString('hh:mm:ss tt')
                }
                "*PT*" {
                    $PatchTuesdayCurrentMonth = Get-CCIPatchTuesday -Month (Get-Date).Month -Year (Get-Date).Year
                    if ((Get-Date) -le $PatchTuesdayCurrentMonth) {
                        $PatchStartDate = $PatchTuesdayCurrentMonth.ToString('MM-dd-yyyy')
                    } else {
                        $NextMonthDate = (Get-Date).AddMonths(1)
                        $PatchStartDate = (Get-CCIPatchTuesday -Month $NextMonthDate.Month -Year $NextMonthDate.Year).ToString('MM-dd-yyyy')
                    }
                    $MWDetails = Get-CMMaintenanceWindow -CollectionName $MWCollection
                    $PatchStartTime = $MWDetails.StartTime.ToString('hh:mm:ss tt')
                }
                default {
                    $MWDetails = Get-CMMaintenanceWindow -CollectionName $MWCollection
                    $PatchStartDate = $MWDetails.StartTime.ToString('MM-dd-yyyy')
                    $PatchStartTime = $MWDetails.StartTime.ToString('hh:mm:ss tt')
                }
            }
            foreach ($CollectionMember in $CollectionMembers) {
                $MachineRecord = [PSCustomObject][ordered]@{
                    Name = $CollectionMember.Name
                    Collection = $MWCollection
                    PatchStartDate = $PatchStartDate
                    PatchStartTime = $PatchStartTime
                }
                $MachinesWithMW += $MachineRecord
            }
        }
    }

    end {
        Set-Location $CurrentLocation | Out-Null
        if ($ShowUnassignedActiveMachines) {
            # Compare the list of devices that are members of an update group to the list of all active devices and report the difference.
            $UnassignedServers = Compare-Object -ReferenceObject $ActiveDevices -DifferenceObject $MachinesWithMW -Property Name | Where-Object {
                $_.sideindicator -eq "<="
            } | Select-Object -ExpandProperty Name
            if ($UnassignedServers) {
                # Recommend best patch group for the unassigned servers based on current member counts.
                $PatchGroupCounts = $MachinesWithMW | Where-Object { $_.collection -notmatch "Exchange" } | Group-Object -Property Collection | Sort-Object -Property Count
                $RecommendedGroup = $PatchGroupCounts | Select-Object -First 1
                Write-Verbose "Detected $($UnassignedServers.Count) servers not assigned to a patch collection:"
                $UnassignedServers | ForEach-Object { Write-Verbose "$_" }
                Write-Verbose "Recommended that they join '$($RecommendedGroup.Name)' since it only has $($RecommendedGroup.Count) members"
                $UnassignedServerSuggestions = @()
                foreach ($UnassignedServer in $UnassignedServers) {
                    $UnassignedServerSuggestions += [PSCustomObject]@{
                        ServerName = $UnassignedServer
                        SuggestedPatchCollection = $RecommendedGroup.Name
                        CommandToAdd = "Set-Location -Path $($SiteCode): ; Add-CMDeviceCollectionDirectMembershipRule -CollectionName `"$($RecommendedGroup.Name)`" -Resource `"$UnassignedServer`""
                    }
                }
                return $UnassignedServerSuggestions
            } else {
                Write-Output "There are currently no servers unassigned to a patch collection."
            }
        } else {
            $MachinesWithMW = $MachinesWithMW | Sort-Object -Property PatchStartDate,PatchStartTime,Name
            return $MachinesWithMW
        }
    }
}

Export-ModuleMember -Function 'Get-CCISCCMPatchSchedule'

#endregion Get-CCISCCMPatchSchedule
