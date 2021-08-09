<#
.SYNOPSIS
    Script for automation parralel job to update all computers with latest Firefox.

.DESCRIPTION
    Every time the script is run, it will ping all selected computers, those that are online will be checked for Firefox presence.
    If the Firefox is installed, it will be checked for the version. If not installed, will be skipped.
    If the Firefox version is outdated, the latest will be downloaded and installed. If latest is already installed, will be skipped.
    Everything will be reported in a log at the end, alongside summary. Logs are kept for 30 days.
    The default number of parallel computuers that are processed is defined in $chunk. 15 seems to be optimal for network and cpu load.

.OUTPUTS
    The script will output a .txt file with the date and time of the run in its name at a location specified in LogPath parameter.
#>

$StartTime = (Get-Date)

#Define the number of days to keep the logs
$LogHistoryDays = "-30"

#Define the number of parallel jos
[int]$chunk = 15

#Define messages for the log
$Message_Offline = "OFFLINE"
$Message_NotInstalled = "Not installed"
$Message_LatestInstalled = "Latest installed"
$Message_UpdateSuccess = "Updated successfully"
$Message_UpdateFail = "Update failed"
$Message_LocalPC = "Skipping local PC"
$Message_ErrorRemote = "Error invoking remote command"

#Get all Windows user computers
$Computers = Get-ADComputer -filter * | 
Where-Object {
    $_.Name -like "COMPUTER*" -or $_.Name -like "LAPTOP*"
} | Select-Object -ExpandProperty Name

#Create new log file
$LogFolder = "\\NETWORK\Logs\Firefox"
$LogName = "FF_" + (Get-Date -Format dd-MM-yyyy_HH-mm-ss) + ".txt"
$LogPath = "$LogFolder\$LogName"
New-Item $LogPath -ItemType File | Out-Null

#Ping as job and separate online from offline computers immediately
$PingComputers = $Computers | ForEach-Object { Test-Connection -ComputerName $_ -Count 1 -AsJob } | Get-Job | Receive-Job -Wait
$OnlineComputers = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
$OfflineComputers = $PingComputers | Where-Object { $_.StatusCode -ne 0 } | Select-Object -ExpandProperty Address
Get-Job | Remove-Job

#Get the latest version
$FFLatestLink = "https://product-details.mozilla.org/1.0/firefox_versions.json" 
$FFLatest = (Invoke-WebRequest $FFLatestLink -UseBasicParsing | ConvertFrom-Json).LATEST_FIREFOX_VERSION

#Split array of OnlineComputers into chunks to be processed in parallel
$z = for ($i = 0; $i -lt $OnlineComputers.length; $i += $chunk) { , ($OnlineComputers[$i .. ($i + ($chunk - 1))]) }
[int]$groups = $i / $chunk
for ($n = 0; $n -lt $groups; $n++ ) {
    if ($chunk -ge $OnlineComputers.Length) {
        $array = $z
    }
    else {
        $array = $z[$n]
    }

    $Arguments = @(
        $Message_Offline,
        $Message_NotInstalled,
        $Message_LatestInstalled,
        $Message_UpdateSuccess,
        $Message_UpdateFail,
        $Message_LocalPC,
        $Message_ErrorRemote,
        $LogPath,
        $FFLatest
    )

    $Jobs = $array | ForEach-Object {
        Start-Job -ArgumentList $Arguments -ScriptBlock {
            param (
                $Message_Offline,
                $Message_NotInstalled,
                $Message_LatestInstalled,
                $Message_UpdateSuccess,
                $Message_UpdateFail,
                $Message_LocalPC,
                $Message_ErrorRemote,
                $LogPath,
                $FFLatest
            )
            $computer = $using:_

            #Begin loop for online computers
            if ((Test-Connection -Computer $Computer -Quiet -Count 1) -and (Test-Path "\\$Computer\c$")) {
                #Check x64 or x86 version of the installation
                $FF64Path = "\\$computer\C$\Program Files\Mozilla Firefox\firefox.exe"
                $FF32Path = "\\$computer\C$\Program Files (x86)\Mozilla Firefox\firefox.exe"
                $FF64 = Test-Path -Path $FF64Path
                $FF32 = Test-Path -Path $FF32Path
                if ($FF64 -eq $True) {
                    $FFinstalled = (Get-Item $FF64Path).VersionInfo.Fileversion.ToString()
                    $FFPath = $FF64Path
                }
                else {
                    if ($FF32 -eq $True) {
                        $FFInstalled = (Get-Item $FF32Path).VersionInfo.Fileversion.ToString()
                        $FFPath = $FF32Path
                    }
                    else {
                        $FFinstalled = $False
                    }
                }

                if ($FFInstalled -eq $false) {
                    Write-Output "$Computer - $Message_NotInstalled"
                }
                else {
                    If ($FFInstalled -match $FFLatest) {
                        Write-Output "$Computer - $Message_LatestInstalled"
                    }
                    else {
                        try {
                            Invoke-Command -Computer $computer -Scriptblock {
                                #Download to remote pc
                                $source = "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64&lang=en-US"
                                $location = "C:\Firefox.exe"
                                Invoke-WebRequest $source -OutFile $location
                                Unblock-File -Path $location
                                #Kill process
                                Invoke-CimMethod -Query 'select * from Win32_Process where name like "firefox%"' -MethodName "Terminate" | Out-Null
                                #Install Firefox
                                Start-Process -FilePath $location -ArgumentList "/s" -Wait -PassThru | Out-Null
                                Remove-Item -Path $location -Force
                            } -ErrorAction Stop
                        }
                        catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
                            Write-Output "$Computer - $Message_ErrorRemote"
                            Continue
                        }
                        $FFCheckVersion = (Get-Item $FFPath).VersionInfo.Fileversion.ToString()
                        if ($FFCheckVersion -match $FFLatest) {
                            Write-Output "$Computer - $Message_UpdateSuccess"
                        }
                        else {
                            Write-Output "$Computer - $Message_UpdateFail"
                        }
                    }
                }
            }
            #In case computer went offline between the inital check and the job execution
            else {
                Write-Output "$Computer - $Message_Offline"
            }
        }
    }
    $Jobs | Receive-Job -Wait | Out-File $LogPath -Append
    Get-Job | Remove-Job
}

#List offline computers in log
$OfflineComputers = $OfflineComputers | ForEach-Object { $_ + " - " + $Message_Offline }
$OfflineComputers | Out-File -FilePath $LogPath -Append

#Local PC Skipped
Add-Content -Value "$env:COMPUTERNAME - $Message_LocalPC" -Path $LogPath -Force

#Sort log by computer name
Get-Content $LogPath | Sort-Object | Set-Content $LogPath

#Report at the end of the outcome
$GetLogFile = Get-Content $LogPath
[int]$Delay = 100

$TotalPCsChecked = $Computers.Count
Add-Content -Value "`nTotal Number of Computers Checked: $TotalPCsChecked" -Path $LogPath -Force
Start-Sleep -Milliseconds $Delay
$Report_Offline = ($GetLogFile | Select-String -Pattern $Message_Offline).length
Add-Content -Value "Total Offline: $Report_Offline" -Path $LogPath -Force
Start-Sleep -Milliseconds $Delay
$Report_NotInstalled = ($GetLogFile | Select-String -Pattern $Message_NotInstalled).length
Add-Content -Value "Total Not Installed: $Report_NotInstalled" -Path $LogPath -Force
Start-Sleep -Milliseconds $Delay
$Report_LatestInstalled = ($GetLogFile | Select-String -Pattern $Message_LatestInstalled).length
Add-Content -Value "Total Latest Installed: $Report_LatestInstalled" -Path $LogPath -Force
Start-Sleep -Milliseconds $Delay
$Report_UpdateSuccess = ($GetLogFile | Select-String -Pattern $Message_UpdateSuccess).length
Add-Content -Value "Total Updated Successfully: $Report_UpdateSuccess" -Path $LogPath -Force
Start-Sleep -Milliseconds $Delay
$Report_UpdateFail = ($GetLogFile | Select-String -Pattern $Message_UpdateFail).length
Add-Content -Value "Total Update Failures: $Report_UpdateFail" -Path $LogPath -Force
Start-Sleep -Milliseconds $Delay
$Report_ErrorRemote = ($GetLogFile | Select-String -Pattern $Message_ErrorRemote).length
Add-Content -Value "Total Errors Remoting: $Report_ErrorRemote" -Path $LogPath -Force

#Clean files older than
Get-ChildItem $LogFolder | Where-Object { $_.LastWriteTime -lt ((Get-Date).AddDays($LogHistoryDays)) } | Remove-Item -Force

#Measure execution time
$EndTime = (Get-Date)
$ScriptDuration = New-TimeSpan -Start $StartTime -End $EndTime
Add-Content -Value "`nTotal Execution Time: $($ScriptDuration.Minutes)m:$($ScriptDuration.Seconds)s" -Path $LogPath -Force