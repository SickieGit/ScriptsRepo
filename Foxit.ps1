<#
----------------------------------------------------------------
v 1.1 changelog 14.05.2020.
-implemented automatic download of the installation file, GUI browser removed
-downloaded .exe file name now includes build version to differentiate itself
----------------------------------------------------------------

USER MUST DEFINE:
----------------------------------------------------------------
$Computers (a list, default is from network)
----------------------------------------------------------------
Functions:

Invoke-Check (quick check, looking only for Foxit versions)
Get-Foxit
Copy-Foxit
Install-Foxit
Update-Foxit
Uninstall-Foxit

Logic:
0 Select quick of full check
    0.1 Quick check [exit]
    0.2 Full check

1 Computer turned on
    1.1 Foxit isn't installed
        1.1.1 Install Foxit
        1.1.2 Don't install Foxit
    1.2 Foxit is installed
        1.2.1 Latest version installed
                1.2.1.1 Uninstall Foxit
                1.2.1.2 Leave latest version
        1.2.2 Older version installed
                1.2.2.1 Update Foxit
                1.2.2.2 Uninstall Foxit
2 Computer turned off
3 End check [exit]
#>

<#COMMANDS FOR .msi INSTALLATION/UNINSTALLATION
& cmd /c "msiexec.exe /i c:\Foxit.msi" /qn ADDLOCAL="FX_PDFVIEWER"
& cmd /c "msiexec.exe /x c:\Foxit.msi" /qn CLEAN="1"
#>

Remove-Variable * -ErrorAction SilentlyContinue; Remove-Module *; $error.Clear()
Add-Type -AssemblyName System.Windows.Forms
$Computers = Get-Content -Path "\\NETWORK\Computers.txt"

#Online check for the latest version (top h3 @ 0 position)
$link = Invoke-WebRequest https://www.foxitsoftware.com/pdf-reader/version-history.php
$FoxitLatest = $link.ParsedHtml.body.getElementsByTagName("h3")[0].innerHTML.substring(13)

$exename = "Foxit_$FoxitLatest.exe"
$location = "C:/$exename"
$source = "https://www.foxitsoftware.com/downloads/latest.php?product=Foxit-Reader&platform=Windows&package_type=exe&language=&version=$FoxitLatest"

function Invoke-Check {
    ForEach ($Computer in $Computers) {
        if ($Computer -eq $Env:COMPUTERNAME) {
            Write-Host "$env:COMPUTERNAME skipping local PC" -ForegroundColor Red
        }
        Else {
            if (Test-Connection -Computer $Computer -Quiet -Count 1) {
                try { $FoxitNewInstalled = (Get-WmiObject -ErrorAction Stop -ComputerName $Computer -Class CIM_DataFile -Filter "Name='C:\\Program Files (x86)\\Foxit Software\\Foxit Reader\\FoxitReader.exe'" | Select-Object Version).Version }
                catch [System.Runtime.InteropServices.COMException] {
                    Write-Host "$Computer is ON, but error when getting info" -ForegroundColor Yellow
                    Continue
                }
                if ($null -ne $FoxitNewInstalled) {
                    if ($FoxitNewInstalled -match $FoxitLatest) {
                        Write-Host "$Computer Foxit latest version $FoxitNewInstalled" -ForegroundColor Green
                    }
                    if ($FoxitNewInstalled -notmatch $FoxitLatest) {
                        Write-Host "$Computer Foxit older version $FoxitNewInstalled" -ForegroundColor Yellow
                    }
                }
                if ($null -eq $FoxitNewInstalled) {
                    Write-Host "$Computer Foxit not found"
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
    Invoke-Check
    Exit
}
#0.2 Full check
if ($question0 -eq "2") {
    Write-Host "Full check selected" -ForegroundColor green
    <#GUI File Browser-not used anymore as of v1.1
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter           = 'Text files (*.exe)|*.exe'
        Title            = 'Select installation file'
    }
    $null = $FileBrowser.ShowDialog()
    $location = $FileBrowser.FileName
    $exename = $FileBrowser.SafeFileName
    Write-Host "`nYou have chosen $exename as your installation file" -ForegroundColor Magenta
    Unblock-File -Path "$location"#>
}

# Begin loop for each listed computer
ForEach ($Computer in $Computers) {

    #Define all functions
    function Get-Foxit {
        $Test = Test-Path -Path $location -PathType Leaf
        if ($test -ne $true) {
            Write-Host "Downloading the latest Foxit locally..."
            Invoke-WebRequest $source -OutFile $location
            Unblock-File -Path $location
            Write-Host "Latest Foxit $FoxitLatest successfully downloaded" -ForegroundColor Green
        }
        else {
            Write-Host "Found latest downloaded Foxit installation" -ForegroundColor Green
        }
    }
    function Copy-Foxit {
        Get-Foxit
        Write-Host "Copying the latest Foxit installation..."
        Start-BitsTransfer -Source $location -Destination \\$Computer\c$ -Description "Foxit Installation" -DisplayName "$exename"
        if ((Test-Path -Path "\\$computer\C$\$exename" -PathType Leaf) -eq $true) {
            Write-Host "Copying complete" -ForegroundColor Green
        }
        else {
            Write-Host "Copying failed" -ForegroundColor Red
        }
    }

    function Install-Foxit {
        Write-Host "Installing the latest Foxit, can take up to 90 sec..."
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            Stop-Process -Name FoxitReader -ErrorAction SilentlyContinue
            $InstallArg = ' /COMPONENTS="FX_PDFVIEWER" /TASKS="setDefaultReader" /S'
            $InstallCommand = "C:/$using:exename" + $InstallArg
            & cmd /c $InstallCommand -wait
            if ((Test-Path -Path "C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe" -PathType Leaf) -eq $true) {
                Write-Host "Installation complete" -ForegroundColor Green
            }
            else {
                Write-Host "Installation failed" -ForegroundColor Red
            }
            Remove-Item "C:\$using:exename" -Force
        }
    }

    function Update-Foxit {
        Write-Host "Updating Foxit to the latest version, can take up to 90 sec..."
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            Stop-Process -Name FoxitReader -ErrorAction SilentlyContinue
            $UpdateArg = ' /COMPONENTS="FX_PDFVIEWER" /TASKS="setDefaultReader" /force /S'
            $UpdateCommand = "C:/$using:exename" + $UpdateArg
            & cmd /c $UpdateCommand -wait
            Remove-Item "C:\$using:exename" -Force
        }
        $FoxitUpdated = (Get-Item -Path "\\$computer\C$\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe").VersionInfo.FileVersion
        if ($FoxitUpdated -eq $FoxitLatest) {
            Write-Host "Updating complete" -ForegroundColor Green
        }
        else {
            Write-Host "Updating failed" -ForegroundColor Red
        }
    }

    function Uninstall-Foxit {
        Write-Host "Uninstalling Foxit $FoxitInstalled..."
        Invoke-Command -ComputerName $Computer -ScriptBlock {
            Stop-Process -Name FoxitReader -ErrorAction SilentlyContinue
            #there can be multiple uninsXXX.exe files, we use this to select the latest valid
            $UninstallExe = Get-ChildItem -Path "C:\Program Files (x86)\Foxit Software\Foxit Reader\" -Name -Include unins*.exe | Sort-Object -Descending | Select-Object -First 1
            $UninstallArg = " /clean /verysilent"
            $UninstallPath = 'C:\Program Files (x86)\Foxit Software\Foxit Reader\'
            $command = '"' + $UninstallPath + $UninstallExe + '"' + $UninstallArg
            & cmd /c $command -wait
            Write-Host "Done, deleting leftovers..."
            $Leftover = "C:\Program Files (x86)\Foxit Software"
            #It needs 8 sec cooldown to delete all of the files
            Start-Sleep -Seconds 8
            & cmd /c rd /s /q $Leftover
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
            try { $FoxitInstalled = (Get-WmiObject -ErrorAction Stop -ComputerName $Computer -Class CIM_DataFile -Filter "Name='C:\\Program Files (x86)\\Foxit Software\\Foxit Reader\\FoxitReader.exe'" | Select-Object Version).Version }
            catch [System.Runtime.InteropServices.COMException] {
                Write-Host "Error getting info from this computer" -ForegroundColor Red
                Continue
            }

            #1.1 Foxit doesn't exist
            if ($null -eq $FoxitInstalled) {
                Write-Host "Foxit not found" -ForegroundColor Yellow
                
                #Ask to install
                while (($question1 = Read-Host "Install Foxit? [1]-Yes [2]-No ") -notmatch '^1$|^2$') {
                    write-host "Incorrect entry, try again." -ForegroundColor Red 
                }
                    
                #1.1.1 User accepts installation
                if ($question1 -eq "1") {
                    Copy-Foxit
                    Install-Foxit
                }
                                      
                #1.1.2 User denies installation
                if ($question1 -eq "2") { 
                    Write-Host "Skipped Foxit installation" -ForegroundColor Yellow
                }

            }
            
            #1.2 If Foxit already exists
            Else {
                #1.2.1 Latest Foxit version already installed
                If ($FoxitInstalled -match $FoxitLatest) {

                    Write-Host "The latest Foxit $FoxitInstalled is installed" -ForegroundColor Green
                    while (($question7 = Read-Host "Uninstall? [1]-Yes [2]-No") -notmatch '^1$|^2$') {
                        Write-Host "Incorrect entry, try again." -ForegroundColor Red 
                    }
                    
                    #1.2.1.1 User chooses to uninstall Foxit
                    if ($question7 -eq "1") {
                        Uninstall-Foxit
                    }

                    #1.2.1.2 User chooses to leave the latest version
                    if ($question7 -eq "2") {
                        Write-Host "Latest Foxit $FoxitInstalled remained on $Computer" -ForegroundColor Green
                    }
                    
                }
      
                #1.2.2 Older Foxit version installed
                Else {         
                    Write-Host "Choose what to do with old Foxit $FoxitInstalled" -ForegroundColor Yellow
                    while (($question3 = Read-Host "[1]-Update [2]-Uninstall") -notmatch '^1$|^2$') {
                        write-host "Incorrect entry, try again." -ForegroundColor Red 
                    }
                    
                    #1.2.2.1 User chooses to update Foxit
                    if ($question3 -eq "1") {
                        Copy-Foxit
                        Update-Foxit
                    }
                    #1.2.2.2 User chooses to uninstall Foxit
                    if ($question3 -eq "2") {
                        Uninstall-Foxit
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