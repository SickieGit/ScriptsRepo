function Get-LocalComputerDetails {
    <#
  .SYNOPSIS
      Get all of the relevant hardware information about the computer

  .DESCRIPTION
      The script checks the computer for the basic hardware information:
      CPU: Name, cores and threads
      RAM: Total, total in slots, total slots used and type
      MB: Product, Manufacturer, SerialNumber, Version
      Disk: Size, Free space
      OS: Version, boot time and user logged in

  .PARAMETER Computers
      Enter one or more computers to check. Mandatory

  .OUTPUTS
      The script will output the information to the console window.

  .EXAMPLE
      Get-LocalComputerDetails

      Description
      Will retrieve all of the required hardware details for the local computer
  #>
    $CPU_Name = (Get-CimInstance -ClassName Win32_Processor).Name
    $CPU_Cores = (Get-CimInstance -ClassName Win32_Processor | Measure-Object -Property NumberOfCores -Sum).Sum
    $CPU_Threads = (Get-CimInstance Win32_ComputerSystem).NumberOfLogicalProcessors

    $RAM_Total = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1gb
    $RAM_SlotsTotal = (Get-CimInstance Win32_PhysicalMemoryArray).MemoryDevices
    $RAM_SlotsUsed = (Get-CimINstance Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Count
    $RAM_GetType = Get-CimInstance Win32_PhysicalMemory | Select-Object SMBIOSMemoryType | Select-Object -First 1
    if ($RAM_GetType -eq 21) { $RAM_Type = "DDR2" }
    if ($RAM_GetType -eq 24) { $RAM_Type = "DDR3" }
    if ($RAM_GetType -eq 26) { $RAM_Type = "DDR4" }

    $MB_Info = Get-CimInstance win32_baseboard | Format-List Product, Manufacturer, SerialNumber, Version

    $Disk_Info = Get-CimInstance Win32_LogicalDisk -Filter DriveType=3 | Select-Object DeviceID, @{'Name' = 'Size (GB)'; 'Expression' = { [string]::Format('{0:N0}', [math]::truncate($_.size / 1GB)) } }, @{'Name' = 'Freespace (GB)'; 'Expression' = { [string]::Format('{0:N0}', [math]::truncate($_.freespace / 1GB)) } }

    $OS_Version = (Get-CimInstance Win32_OperatingSystem).Version 
    $OS_BootTime = (Get-CimInstance -ClassName win32_operatingsystem).lastbootuptime
    $OS_LoggedIn = (Get-Process Explorer -IncludeUsername | Where-Object { $_.Username -notlike "*SYSTEM" }).Username

    Write-Host "$Computer" -ForegroundColor Green
    Write-Host "`nCPU INFO:" -ForegroundColor Yellow
    "CPU: $CPU_Name
Cores: $CPU_Cores
Threads: $CPU_Threads"
    Write-Host "`nRAM INFO:" -ForegroundColor Yellow
    "RAM: $RAM_Total GB
RAM Slots Total: $RAM_SlotsTotal
RAM Slots Used: $RAM_SlotsUsed
Type: $RAM_Type"
    Write-Host "`nMB INFO:" -ForegroundColor Yellow
    Write-Output $MB_Info
    Write-Host "DISK INFO:" -ForegroundColor Yellow
    Write-Output $Disk_Info
    "`n1"
    Write-Host "SYS INFO" -ForegroundColor Yellow
    "`nOS Version: $OS_Version
Boot Time: $OS_BootTime
Users Logged In: $OS_LoggedIn"
}

function Get-RemoteComputerDetails {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$Computers
    )
    
    foreach ($computer in $computers) {
        Invoke-Command -ComputerName $computer -ScriptBlock {
            Get-LocalComputerDetails
        }
    }
}