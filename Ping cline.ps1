function Test-Ping {
    <#
  .SYNOPSIS
      Ping the specified computers in parallel.

  .DESCRIPTION
      The scirpt will ping the specified computers (from .txt file or from AD directly) and separate online from offline.
      It will skip the local computer if located in the list.
      At the end, it will output to console window the online from offline.

  .PARAMETER Computers
      A parameter that can accept multiple computer names.

  .OUTPUTS
      The script will output to console window online, then offline computers

  .EXAMPLE
      Test-Ping -Computers $ComputerList

      Description
      Will ping all of the computers specified in the varaible $ComputerList and separate online from offline ones.
  #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$Computers
    )
    $PingComputers = $ComputerList | ForEach-Object { Test-Connection -ComputerName $_ -Count 1 -AsJob } | Get-Job | Receive-Job -Wait
    $OnlineComputers = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
    $OfflineComputers = $PingComputers | Where-Object { $_.StatusCode -ne 0 } | Select-Object -ExpandProperty Address
    Get-Job | Remove-Job
    Write-Host "ONLINE" -ForegroundColor Green
    $OnlineComputers
    Write-Host "OFLINE" -ForegroundColor Red
    $OfflineComputers
}

#Starting loop, to force user to select one of two choices. Entering anything other than A or T will force the function to run again.
while (($confirmation = Read-Host "Ping from AD or .txt? [1]-AD [2}-.txt") -notmatch '^1$|^2$') {
    write-host "Incorrect entry, try again." -ForegroundColor Red 
}

#If the user chooses to run query from AD
if ($confirmation -eq "1") {
    #Ask user to enter search criteria. Entering only * will search entire AD.
    $ComputerNames = Read-Host -Prompt 'Enter AD search criteria'
    #Ping from AD according to search criteria
    $ComputerList = Get-ADComputer -Filter "Name -like '$ComputerNames'" | Select-Object -ExpandProperty name
    Test-Ping -Computers $ComputerList
}

#If the user choses to run query from a list
if ($confirmation -eq "2") {
    $File = Read-Host -Prompt 'Enter .txt file location, i.e. C:/computers.txt)'
    #Ping from list
    $ComputerList = Get-Content "$File"
    Test-Ping -Computers $ComputerList
}