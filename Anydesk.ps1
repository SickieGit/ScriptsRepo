<#
USER MUST DEFINE:
----------------------------------------------------------------
$Computers (a list in .txt)
----------------------------------------------------------------
Functions:

Invoke-Check (quick check, looking only for AnyDesk versions)
Copy-AnyDesk
Install-AnyDesk
Update-AnyDesk
Set-AnyDeskPass
Uninstall-AnyDesk

Logic:
0 Select quick of full check
    0.1 Quick check [exit]
    0.2 Full check

1 Computer turned on
    1.1 AnyDesk isn't installed
        1.1.1 Install AnyDesk (can set pass here)
        1.1.2 Don't install AnyDesk
    1.2 AnyDesk is installed
        1.2.1 Latest version installed
                1.2.1.1 Uninstall AnyDesk
                1.2.1.2 Leave latest version (can set pass here)
        1.2.2 Older version installed
                1.2.2.1 Update AnyDesk (can set pass here)
                1.2.2.2 Uninstall AnyDesk
                1.2.2.3 Leave (can set pass here)
2 Computer turned off
3 End check [exit]
#>

Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear()
$List = Read-Host "Enter the path to the .txt list of computer names"
$Computers = Get-Content -Path $List
$source = "https://download.anydesk.com/AnyDesk-CM.exe"
$AnyDeskLatest = (Get-Item -Path C:\AnyDesk-CM.exe -ErrorAction SilentlyContinue).VersionInfo.FileVersion
$exename = "AnyDesk-CM.exe"
$location = "C:/$exename"

function Get-AnyDesk {
    $Test = Test-Path -Path "C:/$exename" -PathType Leaf
    if ($Test -eq $True) {
        Remove-Item -Path "C:/$exename" -Force
    }
    Write-Host "Downloading latest AnyDesk..."
    Invoke-WebRequest $source -OutFile $location
    Unblock-File -Path "C:/$exename"
    Write-Host "Download complete" -ForegroundColor Green
}
function Invoke-Check {
    ForEach ($Computer in $Computers) {
        if ($Computer -eq $Env:COMPUTERNAME) {
            Write-Host "$env:COMPUTERNAME skipping local PC" -ForegroundColor Red
        }
        Else {
            if (Test-Connection -Computer $Computer -Quiet -Count 1) {
                try { $AnyDeskNewInstalled = (Get-WmiObject -ErrorAction Stop -ComputerName $Computer -Class CIM_DataFile -Filter "Name='C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe'" | Select-Object Version).Version }
                catch [System.Runtime.InteropServices.COMException] {
                    Write-Host "$Computer is ON, but error when getting info" -ForegroundColor Yellow
                    Continue
                }
                if ($null -ne $AnyDeskNewInstalled) {
                    if ($AnyDeskNewInstalled -match $AnyDeskLatest) {
                        Write-Host "$Computer AnyDesk latest version $AnyDeskNewInstalled" -ForegroundColor Green
                    }
                    if ($AnyDeskNewInstalled -notmatch $AnyDeskLatest) {
                        Write-Host "$Computer AnyDesk older version $AnyDeskNewInstalled" -ForegroundColor Yellow
                    }
                }
                if ($null -eq $AnyDeskNewInstalled) {
                    Write-Host "$Computer AnyDesk not found"
                }
            }
            Else {
                Write-Host "$Computer is OFF" -ForegroundColor Red
            }
        }
    }
}

#0 quick check will only run through all of the computers to see the version
while (($question0 = Read-Host "Choose check type [1]-Quick [2]-Full ") -notmatch '^1$|^2$') {
    write-host "Incorrect entry, try again." -ForegroundColor Red 
}
#0.1 Quick check
if ($question0 -eq "1") {
    Write-Host "Quick check selected" -ForegroundColor Green
    Get-AnyDesk
    Invoke-Check
    Exit
}
#0.2 Full check
if ($question0 -eq "2") {
    Write-Host "Full check selected" -ForegroundColor green
    Get-AnyDesk
}

# Begin loop for each listed computer
ForEach ($Computer in $Computers) {

    #Define all functions
    function Copy-AnyDesk {
        Write-Host "Copying the latest AnyDesk installation..."
        Start-BitsTransfer -Source $location -Destination \\$Computer\c$ -Description "AnyDesk Installation" -DisplayName "$exename"
        #Copy-Item -Path $location -Destination \\$Computer\c$ -Force -Recurse
        Write-Host "Copying complete" -ForegroundColor Green
    }

    function Install-AnyDesk {
        Write-Host "Installing the latest AnyDesk..."
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            Start-Process -FilePath "C:\$using:exename" -ArgumentList '--install "C:\Program Files (x86)\AnyDesk" --start-with-win --create-shortcuts --create-desktop-icon --silent' -Wait
            if ((Test-Path -Path "C:\Program Files (x86)\Anydesk\AnyDesk.exe" -PathType Leaf) -eq $true) {
                Write-Host "Installation complete" -ForegroundColor Green
            }
            else {
                Write-Host "Installation failed" -ForegroundColor Red
            }
            Remove-Item "C:\AnyDesk-CM.exe" -Force
        }
    }

    function Update-AnyDesk {
        Write-Host "Updating AnyDesk to the latest version..."
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            Start-Process -FilePath "C:\program files (x86)\AnyDesk\AnyDesk.exe" -ArgumentList '--silent --remove' -Wait
            Start-Process -FilePath "C:\$using:exename" -ArgumentList '--install "C:\Program Files (x86)\AnyDesk" --start-with-win --create-shortcuts --create-desktop-icon --silent' -Wait
            Remove-Item "C:\AnyDesk-CM.exe" -Force
        }
        $AnyDeskUpdated = (Get-WmiObject -ErrorAction Stop -ComputerName $Computer -Class CIM_DataFile -Filter "Name='C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe'" | Select-Object Version).Version 
        if ($AnyDeskUpdated -eq $AnyDeskLatest) {
            Write-Host "Updating complete" -ForegroundColor Green
        }
        else {
            Write-Host "Installation failed" -ForegroundColor Red
        }
    }
    function Set-AnyDeskPass {
        Write-Host "Set up password for unattended access?"
        while (($question8 = Read-Host "[1]-Yes [2]-No") -notmatch '^1$|^2$') {
            write-host "Incorrect entry, try again." -ForegroundColor Red 
        }
        if ($question8 -eq "1") {
            $s = New-PSSession -ComputerName $computer
            $Password = Read-Host "Create a Password"
            Invoke-Command -Session $s -ScriptBlock {
                $CheckID = @'
@echo off & for /f "delims=" %i in ('"C:\Program Files (x86)\AnyDesk\AnyDesk.exe" --get-id') do set CID=%i & call echo %CID%
'@
                &cmd /c $CheckID | Out-File -FilePath C:/output.txt
                $SetPassword1 = "echo $using:Password"
                $SetPassword2 = ' | "C:\Program Files (x86)\AnyDesk\AnyDesk.exe" --set-password'
                $SetPassword = $SetPassword1 + $SetPassword2
                &cmd /c $SetPassword
            }
            Get-PSSession | Remove-PSSession
            $AnyDeskID = Get-Content -Path \\$computer\c$\output.txt
            Remove-Item -Path \\$computer\c$\output.txt -Force
            Write-Host "AnyDesk ID = $AnyDeskID" -ForegroundColor Green
            Write-Host "AnyDesk password = $Password" -ForegroundColor Green
        }
        if ($question8 -eq "2") {
            Write-Host "Password won't be set"
        }
    }

    function Uninstall-AnyDesk {
        $ErrorActionPreference = "SilentlyContinue"
        Write-Host "Uninstalling AnyDesk $AnyDeskInstalled..."
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            Start-Process -FilePath "C:\program files (x86)\AnyDesk\AnyDesk.exe" -ArgumentList '--silent --remove' -Wait
            Remove-Item -Path "C:\program files (x86)\AnyDesk\" -force -recurse
        }
        Write-Host "Uninstallation complete" -ForegroundColor Green
    }
    
    if ($Computer -match $env:COMPUTERNAME | Select-Object) {
        Write-Host "`nSkipping $env:COMPUTERNAME as it is a local computer" -ForegroundColor Red
    }

    Else {

        #1 - If the computer is turned on
        if (Test-Connection -Computer $Computer -Quiet -Count 1) {
            Write-Host "`n$Computer is ON" -ForegroundColor Green
            #Test if there are issues with the computer, if not move on
            try { $AnyDeskInstalled = (Get-WmiObject -ErrorAction Stop -ComputerName $Computer -Class CIM_DataFile -Filter "Name='C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe'" | Select-Object Version).Version }
            catch [System.Runtime.InteropServices.COMException] {
                Write-Host "Error getting info from this computer" -ForegroundColor Red
                Continue
            }
            #1.1 AnyDesk doesn't exist
            if ($null -eq $AnyDeskInstalled) {
                Write-Host "AnyDesk not found" -ForegroundColor Yellow
                
                #Ask to install
                while (($question1 = Read-Host "Install AnyDesk? [1]-Yes [2]-No ") -notmatch '^1$|^2$') {
                    write-host "Incorrect entry, try again." -ForegroundColor Red 
                }
                    
                #1.1.1 User accepts installation
                if ($question1 -eq "1") {
                    Copy-AnyDesk
                    Install-AnyDesk
                    Set-AnyDeskPass
                }
                                      
                #1.1.2 User denies installation
                if ($question1 -eq "2") { 
                    Write-Host "Skipped AnyDesk installation" -ForegroundColor Yellow
                }

            }
            
            #1.2 If AnyDesk already exists
            Else {
                #1.2.1 Latest AnyDesk version already installed
                If ($AnyDeskInstalled -match $AnyDeskLatest) {

                    Write-Host "The latest AnyDesk $AnyDeskInstalled is installed" -ForegroundColor Green
                    while (($question7 = Read-Host "Uninstall? [1]-Yes [2]-No") -notmatch '^1$|^2$') {
                        Write-Host "Incorrect entry, try again." -ForegroundColor Red 
                    }
                    
                    #1.2.1.1 User chooses to uninstall AnyDesk
                    if ($question7 -eq "1") {
                        Uninstall-AnyDesk
                    }

                    #1.2.1.2 User chooses to leave the latest version
                    if ($question7 -eq "2") {
                        Write-Host "Latest AnyDesk $AnyDeskInstalled remained on $Computer" -ForegroundColor Green
                        Set-AnyDeskPass
                    }
                    
                }
      
                #1.2.2 Older AnyDesk version installed
                Else {         
                    Write-Host "Choose what to do with old AnyDesk $AnyDeskInstalled" -ForegroundColor Yellow
                    while (($question3 = Read-Host "[1]-Update [2]-Uninstall [3]-Leave") -notmatch '^1$|^2$|^3$') {
                        write-host "Incorrect entry, try again." -ForegroundColor Red 
                    }
                    
                    #1.2.2.1 User chooses to update AnyDesk
                    if ($question3 -eq "1") {
                        Copy-AnyDesk
                        Update-AnyDesk
                        Set-AnyDeskPass
                    }
                    #1.2.2.2 User chooses to uninstall AnyDesk
                    if ($question3 -eq "2") {
                        Uninstall-AnyDesk
                    }
                    #1.2.2.3 User chooses to leave old AnyDesk
                    if ($question3 -eq "3") {
                        Write-Host "AnyDesk $AnyDeskInstalled remains" -ForegroundColor Yellow
                        Set-AnyDeskPass
                    }
                }
            }
        }
   
        #2 If the computer is turned off
        Else {
            Write-Host "`n$Computer is OFF" -ForegroundColor Red
        }
    }
}
#3 End check for results
Write-Host "`nScript complete, checking results...`n" -ForegroundColor Magenta
Invoke-Check