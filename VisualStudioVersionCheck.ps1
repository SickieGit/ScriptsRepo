<#
---------------------------------------------
v 1.1 changelog
-changed $VS1$Latest online parsing to lookup latest release json file due to unreliable 1st method due to shifting tables
-changed logic for comparing local to online versions to account for possbile preview releases
-minor bug for $VS17E comparison fixed
---------------------------------------------
This script will check the listed computers for the version of the Visual Studio that is installed.
It will look for 2019 and 2017 installation for all three versions - Professional, Enterprise and Community.
Then it will compare that version to the latest stable release version of Visual Studio from the Internet.

Color coding output will look like this:
GREEN - Latest version is installed
YELLOW - Older version is installed
BLUE - Doesn't exist
RED - Computer is off
#>

$ErrorActionPreference = "SilentlyContinue"
$Computers = Get-Content -Path "\\network\Computers.txt"

$VS19Latest = Invoke-WebRequest "https://aka.ms/vs/16/release/channel" | ConvertFrom-Json | Select-Object -ExpandProperty info | Select-Object -ExpandProperty BuildVersion
$VS17Latest = Invoke-WebRequest "https://aka.ms/vs/15/release/channel" | ConvertFrom-Json | Select-Object -ExpandProperty info | Select-Object -ExpandProperty BuildVersion

foreach ($computer in $computers) {

    if (Test-Connection -Computer $computer -Quiet -Count 1) {
        $Owner = Get-ADComputer -Filter 'name -like $computer' -Properties description | Select-Object -ExpandProperty description
        Write-Host "`n$computer - $Owner"
        $VS19 = Test-Path -Path "\\$computer\c$\Program Files (x86)\Microsoft Visual Studio\2019\"
        $VS17 = Test-Path -Path "\\$computer\c$\Program Files (x86)\Microsoft Visual Studio\2017\"

        function Get-VSVersion {
            param (
                $Year, $Version
            )
            (Get-Command "\\$computer\C$\Program Files (x86)\Microsoft Visual Studio\$Year\$Version\Common7\IDE\devenv.exe").FileVersionInfo.ProductVersion
        }

        if ($VS19 -eq $True) {
            $VS19P = Get-VSVersion -Year "2019" -Version "Professional"
            $VS19E = Get-VSVersion -Year "2019" -Version "Enterprise"
            $VS19C = Get-VSVersion -Year "2019" -Version "Community"

            if ($Null -ne $VS19P) {
                if ($VS19P -ge $VS19Latest) {
                    Write-Host "2019 Professional $VS19P" -ForegroundColor Green
                }
                else {
                    Write-Host "2019 Professional $VS19P" -ForegroundColor Yellow
                }
            
            }
            if ($Null -ne $VS19E) {
                if ($VS19E -ge $VS19Latest) {
                    Write-Host "2019 Enterprise $VS19E" -ForegroundColor Green
                }
                else {
                    Write-Host "2019 Enterprise $VS19E" -ForegroundColor Yellow
                }
            }
            if ($Null -ne $VS19C) {
                if ($VS19C -ge $VS19Latest) {
                    Write-Host "2019 Community $VS19C" -ForegroundColor Green
                }
                else {
                    Write-Host "2019 Community $VS19C" -ForegroundColor Yellow
                }
            }
            if ($Null -eq $VS19P -and $Null -eq $VS19E -and $Null -eq $VS19C) {
                Write-Host "2019 doesn't exist" -ForegroundColor Cyan
            }
        }
        if ($VS19 -eq $False) {
            Write-Host "2019 doesn't exist" -ForegroundColor Cyan
        }
        if ($VS17 -eq $True) {
            $VS17P = Get-VSVersion -Year "2017" -Version "Professional"
            $VS17E = Get-VSVersion -Year "2017" -Version "Enterprise"
            $VS17C = Get-VSVersion -Year "2017" -Version "Community"

            if ($Null -ne $VS17P) {
                if ($VS17P -ge $VS17Latest) {
                    Write-Host "2017 Professional $VS17P" -ForegroundColor Green
                }
                else {
                    Write-Host "2017 Professional $VS17P" -ForegroundColor Yellow
                }
            }
            if ($Null -ne $VS17E) {
                if ($VS17E -ge $VS17Latest) {
                    Write-Host "2017 Enterprise $VS17E" -ForegroundColor Green
                }
                else {
                    Write-Host "2017 Enterprise $VS17E" -ForegroundColor Yellow
                }
            }
            if ($Null -ne $VS17C) {
                if ($VS17C -ge $VS17Latest) {
                    Write-Host "2017 Community $VS17C" -ForegroundColor Green
                }
                else {
                    Write-Host "2017 Community $VS17C" -ForegroundColor Yellow
                }
            }
            if ($Null -eq $VS17P -and $Null -eq $VS17E -and $Null -eq $VS17C) {
                Write-Host "2017 doesn't exist" -ForegroundColor Cyan
            }
        }
        if ($VS17 -eq $False) {
            Write-Host "2017 doesn't exist" -ForegroundColor Cyan
        }
    }
    else {
        Write-Host "`n$computer is OFF" -ForegroundColor Red
    }
}