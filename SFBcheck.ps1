<#
.SYNOPSIS
    Check all of the specified computers for the presence of Skype for Business.

.DESCRIPTION
    The script will create the starting .csv mini database if one does not already exist.
    It will then ping all computers and for all that are online, check if SFB exists.
    It will report all of the findings to the .csv in detail, as well write the summary to the console.

.OUTPUTS
    The script will output a csv file at a location specified in LogFilePath parameter.
#>

$LogFilePath = "\\NETWORK\Logs\SFB.csv"

if (!(Test-Path $LogFilePath)) {
    Get-ADComputer -filter * | Where-Object {
        $_.Name -like "COMPUTER*" -or 
        $_.Name -like "LAPTOP*"
    } | Select-Object Name, @{n = 'SFB'; e = { "UNKNOWN" } } |
    Sort-Object -Property Name | 
    Export-Csv -Path $LogFilePath -NoTypeInformation
}

$Computers = Import-Csv -Path $LogFilePath
$UnknownComputers = $Computers | Where-Object { $_.SFB -eq "UNKNOWN" }

$PingComputers = $UnknownComputers | ForEach-Object { 
    Test-Connection -ComputerName $_.Name -Count 1 -AsJob } | 
Get-Job | Receive-Job -Wait
$OnlineComputers = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | 
Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
Get-Job | Remove-Job

foreach ($Computer in $OnlineComputers) {

    $ComputerPath = "\\$computer\c$"
    $ProgramFiles64Path = "\Program Files\Microsoft Office"
    $ProgramFiles86Path = "\Program Files (x86)\Microsoft Office"
    $OfficeC2RPath = "\root\Office16\lync.exe"
    $OfficeMSIPath = "\Office16\lync.exe"
    $RowIndex = [array]::IndexOf($Computers.Name, $Computer)

    $CR2x86 = Test-Path ($ComputerPath + $ProgramFiles86Path + $OfficeC2RPath)
    $MSIx86 = Test-Path ($ComputerPath + $ProgramFiles86Path + $OfficeMSIPath)
    $CR2x64 = Test-Path ($ComputerPath + $ProgramFiles64Path + $OfficeC2RPath)
    $MSIx64 = Test-Path ($ComputerPath + $ProgramFiles64Path + $OfficeMSIPath)

    If ($CR2x86 -eq $true -or $MSIx86 -eq $true -or $CR2x64 -eq $true -or $MSIx64 -eq $true ) {
        $Computers[$RowIndex].SFB = "YES"
    }
    else {
        $Computers[$RowIndex].SFB = "NO"
    }
}
$Computers | Export-Csv -Path $LogFilePath -NoTypeInformation

$City1 = $Computers | Where-Object { $_.Name -like "B*" }
$City1Unknown = $City1 | Where-Object { $_.SFB -eq "UNKNOWN" }
$City1Yes = $City1 | Where-Object { $_.SFB -eq "YES" }
$City1No = $City1 | Where-Object { $_.SFB -eq "NO" }

$City2 = $Computers | Where-Object { $_.Name -like "R*" }
$City2Unknown = $City2 | Where-Object { $_.SFB -eq "UNKNOWN" }
$City2Yes = $City2 | Where-Object { $_.SFB -eq "YES" }
$City2No = $City2 | Where-Object { $_.SFB -eq "NO" }

$City3 = $Computers | Where-Object { $_.Name -like "M*" }
$City3Unknown = $City3 | Where-Object { $_.SFB -eq "UNKNOWN" }
$City3Yes = $City3 | Where-Object { $_.SFB -eq "YES" }
$City3No = $City3 | Where-Object { $_.SFB -eq "NO" }


$Table = @(
    @{LOCATION = "City1"; INSTALLED = [int]$City1Yes.Count; NOTINSTALLED = [int]$City1No.Count; UNCHECKED = [int]$City1Unknown.Count },
    @{LOCATION = "City3"; INSTALLED = [int]$City3Yes.Count; NOTINSTALLED = [int]$City3No.Count; UNCHECKED = [int]$City3Unknown.Count },
    @{LOCATION = "City2"; INSTALLED = [int]$City2Yes.Count; NOTINSTALLED = [int]$City2No.Count; UNCHECKED = [int]$City2Unknown.Count },
    @{LOCATION = "-------"; INSTALLED = "-------"; NOTINSTALLED = "-------"; UNCHECKED = "-------" },
    @{LOCATION       = "TOTAL";
        INSTALLED    = $City1Yes.Count + $City2Yes.Count + $City3Yes.Count; 
        NOTINSTALLED = $City1No.Count + $City2No.Count + $City3No.Count; 
        UNCHECKED    = $City1Unknown.Count + $City2Unknown.Count + $City3Unknown.Count
    }) | ForEach-Object { New-Object object | Add-Member -NotePropertyMembers $_ -PassThru }

$Table | Format-Table -Property LOCATION, INSTALLED, NOTINSTALLED, UNCHECKED

Write-Host "Full report at $LogFilePath"