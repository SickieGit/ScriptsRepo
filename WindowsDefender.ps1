<#
.SYNOPSIS
    This script will perform the Windows Defender ATP Onboarding.

.DESCRIPTION
    The script gets the specified list of target computers, pings those that are online. 
    If computer is online, checks for MsMpEng process to see if the Defender onboarding is already done.
    If not, runs the Microsoft Deployment script on it and checks to confirm that the process is running now.
    Returns all of the findings for each computer to the console

.OUTPUTS
    The script will output all findings to the active console.
#>
$ListOfComputers = "\\network\ComputersDefender.txt"
$computers = Get-Content $ListOfComputers
#Ping all computers from the list, then separate online and offline
$PingComputers = $Computers | ForEach-Object { Test-Connection -ComputerName $_ -Count 1 -AsJob } | Get-Job | Receive-Job -Wait
$OnlineComputers = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
$OfflineComputers = $PingComputers | Where-Object { $_.StatusCode -ne 0 } | Select-Object -ExpandProperty Address

foreach ($computer in $OnlineComputers) {
    #Test to see if there are DNS issues
    if (Test-Path "\\$computer\C$") {
        #Test for process MsMpEng to see if Defender is already deployed, proceed if not
        $TestDefenderProc = Get-Process "MsMpEng" -ComputerName $computer -ErrorAction SilentlyContinue
        if (!$TestDefenderProc) {
            Write-Host "$computer - Trying to deploy remotely..." -ForegroundColor Magenta
            #Copy .cmd script from network to target computer
            $ScriptName = "WinDefenderRemote.cmd"
            $networkPath = "\\network\Windows Defender Install\"
            $ScriptPathnetwork = $networkPath + $ScriptName
            $CopyLocation = "\\$computer\C$\temp\"
            Copy-Item -Path $ScriptPathnetwork -Destination $CopyLocation
            #try to run the script remotely, catch PS remoting issues
            try {
                Invoke-Command -ComputerName $computer -ArgumentList $ScriptName -ScriptBlock {
                    $ScriptPathLocal = "C:\temp\" + $using:ScriptName
                    & cmd /c "$ScriptPathLocal" -wait
                } -ErrorAction Stop
            }
            catch [System.Management.Automation.Remoting.PSRemotingTransportException] {
                Write-Host "$computer is having issues with PS remoting" -ForegroundColor Red
                Continue
            }
            #If the deployment was success, process MsMpEng should be running, test this
            $TestDefenderProcNew = Get-Process "MsMpEng" -ComputerName $computer -ErrorAction SilentlyContinue
            if ($TestDefenderProcNew) {
                Write-Host "$Computer successfully onboarded" -ForegroundColor Green
            }
            else {
                Write-Host "$Computer onboarding failed" -ForegroundColor Red
            }
        }
        else {
            Write-Host "$Computer already has Windows Defender" -ForegroundColor Green
        }
    }
    else {
        Write-Host "$computer is offline" -ForegroundColor Red
    }
}
foreach ($OfflineComputer in $OfflineComputers) {
    Write-Host "$OfflineComputer is offline" -ForegroundColor Red
}