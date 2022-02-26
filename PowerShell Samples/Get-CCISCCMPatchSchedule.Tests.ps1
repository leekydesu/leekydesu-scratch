<#

    Get-CCISCCMPatchSchedule.Tests.ps1

    Get-CCISCCMPatchSchedule returns information pertaining to the patch schedules setup in SCCM

#>

# Remove any pre-existing or already loaded version of the module.
Get-Module CCIModule | Remove-Module

# Import the module to be tested (up two folders)
Import-Module "$PSScriptRoot\..\..\src\CCIModule.psd1"

# Run tests inside the module scope so that internal private commands can be tested and mocked.
InModuleScope CCIModule {

    ################
    ## Unit Tests ##
    ################

    Describe 'public function: Get-CCISCCMPatchSchedule - Returns all AD account attribute changes made with in a given time period.' -Tags 'UnitTest' {
        Context 'Parameter Tests' {

            # Importing the ConfigurationManager module for mocks is difficult, so a subset of necessary dummy functions
            # gets created here instead.

            function Get-CMCollection {
                return $null
            }
            function Get-CMCollectionMember {
                return $null
            }
            function Get-CMMaintenanceWindow {
                return $null
            }
            function Get-CMDevice {
                return $null
            }
            function Get-CMSiteSystemServer {
                return $null
            }
            function Get-CMDeployment {
                return $null
            }

            Mock Get-Module {
                return $true
            }
            Mock Test-Path {
                return $true
            }
            Mock Set-Location {
                return $null
            }
            Mock Get-CMCollection {
                $Counter = 1
                $MockCollections = @()
                while ($Counter -le 5) {
                    switch ($Counter) {
                        1 {
                            $MockCollectionName = "Update Group - Exchange Mock"
                        }
                        2 {
                            $MockCollectionName = "Update Group - PT Mock"
                        }
                        3 {
                            $MockCollectionName = "Update Group - Wed Mock"
                        }
                        4 {
                            $MockCollectionName = "Update Group - Sunday Mock"
                        }
                        5 {
                            $MockCollectionName = "Mock Group - Filter Me Out"
                        }
                        default {
                            $MockCollectionName = "Mock Failure Collection"
                        }
                    }
                    $MockCollection = [PSCustomObject]@{
                        Name = $MockCollectionName
                        CollectionId = "CCI0000" + $Counter.ToString("00")
                    }
                    $MockCollections += $MockCollection
                    $Counter++
                }
                return $MockCollections
            }
            Mock Get-CMCollectionMember {
                $MockMembers = @()
                $UpperLimit = Get-Random -Minimum 2 -Maximum 20
                1..$UpperLimit | ForEach-Object {
                    $MockMember = [PSCustomObject]@{
                        Name = "IU-CCI-MOCK" + $_.ToString("00")
                    }
                    $MockMembers += $MockMember
                }
                return $MockMembers
            }
            Mock Get-CMDeployment {
                $MockPatchTime = [PSCustomObject]@{
                    DeploymentTime = (Get-Date).AddDays(14)
                }
                return $MockPatchTime
            }
            Mock Get-CMMaintenanceWindow {
                $MockPatchDate = (Get-Date).AddDays(14)
                $MockPatchWindow = [PSCustomObject]@{
                    StartTime = $MockPatchDate
                }
                return $MockPatchWindow
            }
            Mock Get-CMSiteSystemServer {
                1..2 | ForEach-Object {
                    [PSCustomObject]@{
                        NetworkOSPath = "\\IU-CCI-MOCKSCCM" + $_ + ".ads.iu.edu"
                    }
                    $MockSiteSystemServers += $MockSiteSystemServer
                }
                return $MockSiteSystemServers
            }
            Mock Get-CMDevice {
                $MockDevices = @()
                1..20 | ForEach-Object {
                    if ($_%2) {
                        $IsActive = 1
                    } else {
                        $IsActive = $null
                    }
                    $MockDevice = [PSCustomObject]@{
                        Name = "IU-CCI-MOCK" + $_.ToString("00")
                        IsActive = $IsActive
                    }
                    $MockDevices += $MockDevice
                }
                $MockDevices += [PSCustomObject]@{
                    Name = "IU-CCI-MOCKSCCM1"
                    IsActive = 1
                }
                $MockDevices += [PSCustomObject]@{
                    Name = "IU-CCI-MOCKSCCM2"
                    IsActive = 1
                }
                return $MockDevices
            }
            Mock Compare-Object {
                $MockUnassigned = @()
                1..2 | ForEach-Object {
                    $MockUnassigned += [PSCustomObject]@{
                        Name = "IU-CCI-REBEL0$_"
                    }
                }
                return $MockUnassigned
            }

            #========================================================================================================

            ###########################################
            ### Test Case : Default parameters only ###
            ###########################################

            It 'Default parameters only.' {

                $PatchSchedule = Get-CCISCCMPatchSchedule -ManagementPointFQDN iu-cci-cmmp01.ads.iu.edu -SiteCode CCI
                $PatchSchedule.Count | Should -BeGreaterOrEqual 3
                Assert-MockCalled -CommandName Get-CMCollectionMember -Times 3 -Scope It -Exactly

            }

            #========================================================================================================

            ###################################
            ### Test Case : Include Sundays ###
            ###################################

            It 'Include Sundays.' {

                $PatchSchedule = Get-CCISCCMPatchSchedule -ManagementPointFQDN iu-cci-cmmp01.ads.iu.edu -SiteCode CCI -IncludeSundays
                $PatchSchedule.Count | Should -BeGreaterOrEqual 4
                Assert-MockCalled -CommandName Get-CMCollectionMember -Times 4 -Scope It -Exactly

            }

            #========================================================================================================

            ###############################################
            ### Test Case : Showing unassigned machines ###
            ###############################################

            It 'Showing unassigned machines.' {

                $UnassignedServerList = Get-CCISCCMPatchSchedule -ManagementPointFQDN iu-cci-cmmp01.ads.iu.edu -SiteCode CCI -ShowUnassignedActiveMachines
                $UnassignedServerList.Count | Should -Be 2

            }

        }
        Context 'Error Tests' {
            #========================================================================================================

            ##################################################
            ### Test Case : Incorrect site code or MP FQDN ###
            ##################################################

            It 'Incorrect site code or MP FQDN.' {
                Mock Get-Module {
                    return $true
                }
                { Get-CCISCCMPatchSchedule -ManagementPointFQDN iu-cci-mocksccm.ads.iu.edu -SiteCode MOO } | Should -Throw "Management Point FQDN or Site Code is incorrect."

            }

            #========================================================================================================

            ##############################################
            ### Test Case : SCCM Console not installed ###
            ##############################################

            It 'SCCM Console not installed.' {
                Mock Get-Module {
                    return $false
                }
                Mock Test-Path {
                    return $false
                }

                { Get-CCISCCMPatchSchedule -ManagementPointFQDN iu-cci-cmmp01.ads.iu.edu -SiteCode CCI } | Should -Throw "SCCM Admin Console is not installed on this machine."

            }

        }
    }

}
