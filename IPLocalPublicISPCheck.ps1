<#
  .SYNOPSIS
      For each of the computer specified get local IP, public IP and ISP.

  .DESCRIPTION
      The script checks if computer is online, if yes gets local, public and ISP, then outputs to the console.
      If the computer is offline, thenk it skips and notes in the console

  .OUTPUTS
      The script will output console text with the info specified.
  #>

#Starting loop, to force user to select one of two choices. Entering anything other than 1 or 2 will force the function to run again.
while (($confirmation = Read-Host "Ping from AD or .txt file? [1]-AD [2]-.txt") -notmatch '^1$|^2$') {
    write-host "Incorrect entry, try again." -ForegroundColor Red 
}

#If the user chooses to run query from AD
if ($confirmation -eq "1") {
    #Ask user to enter search criteria. Entering only * will search entire AD.
    $Computers = Read-Host -Prompt 'Enter AD search criteria, for example COMPUTER*'

    #Pull content from AD according to search criteria
    $cmplist = Get-ADComputer -Filter 'Name -like $Computers' | Select-Object -ExpandProperty name
}

#If the user choses to run query from a list
if ($confirmation -eq "2") {
    $File = Read-Host -Prompt 'Enter .txt file location, i.e. C:/computers.txt)'
    #Pull content from list
    $cmplist = Get-Content "$File"
}
$ErrorActionPreference = "SilentlyContinue"
foreach ($cmp in $cmplist) {
    if (Test-Connection -Computer $cmp -Quiet -Count 1) {
        $LocalIp = Get-NetIPAddress -ErrorAction SilentlyContinue -PrefixOrigin Dhcp  -CimSession (New-CimSession -ComputerName $cmp -ErrorAction SilentlyContinue) | Select-Object -ExpandProperty IPAddress
        Get-CimSession -ComputerName $cmp -ErrorAction SilentlyContinue | Remove-CimSession
        $PublicIP = Invoke-Command -ComputerName $cmp -ErrorAction SilentlyContinue -ScriptBlock {
            (Invoke-RestMethod -Uri "http://ipinfo.io")
        } | Select-Object ip, org
        
        Write-Host `nPC name: $cmp -ForegroundColor Yellow
        Write-Host Local IP: $LocalIp -ForegroundColor Green    
        Write-Host Public IP: $PublicIP.ip -ForegroundColor Green
        Write-Host ISP: $PublicIP.org  -ForegroundColor Green
    }
        
    else {
        Write-Host `nPC name: $cmp -ForegroundColor Yellow
        Write-Host Computer offline -ForegroundColor Red
    }
}