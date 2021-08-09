<#
.SYNOPSIS
    Create a separate icon for Firefox with proxy set.

.DESCRIPTION
    The script will check if the computer has Firefox installed.
    If not, it will output the message. If yes, proceed
    It will check if x32 or x64 bit Firefox is installed and set it as a value in variable
    If Firefox has never been used, it will run it and close it to create current profile.
    It will then create a new Firefox profile and edit prefs.js with proxy IP and port
    Finally, a new shortcut will be made to the user's desktop to run the separate instance.

.OUTPUTS
    The script will output to the console window the findings and what has been done.

.NOTES
    The script was made in a way to be compiled in a simple to run .exe file for users to run
    without need for admin elevation.
#>

$FF64Path = "C:\Program Files\Mozilla Firefox\firefox.exe"
$FF32Path = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
$FF64 = Test-Path -Path $FF64Path
$FF32 = Test-Path -Path $FF32Path
if ($FF64 -ne $true -and $FF32 -ne $true) {
    Write-Host "FIREFOX NOT INSTALLED, CALL SERVICE DESK"
    Start-Sleep -Seconds 5
    Break
}
if ($FF64 -eq $true) {
    $FFInstalledPath = $FF64Path 
}
if ($FF32 -eq $true) {
    $FFInstalledPath = $FF32Path
}
Write-Host "Firefox is installed" -ForegroundColor Green
Start-Sleep -Seconds 1
$DefaultProfilePath = "$env:APPDATA\Mozilla\Firefox\Profiles"

function Set-Profile {
    Write-Host "Setting up profile..." -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    Stop-Process -Name firefox -Force -ErrorAction SilentlyContinue
    Start-Process -FilePath $FFInstalledPath -ArgumentList " -CreateProfile Italy" -Wait
    Start-Sleep -Seconds 1
    $ProxyStrings = 'user_pref("network.proxy.backup.ftp", "123.123.123.123");
user_pref("network.proxy.backup.ftp_port", 8080);
user_pref("network.proxy.backup.ssl", "123.123.123.123");
user_pref("network.proxy.backup.ssl_port", 8080);
user_pref("network.proxy.ftp", "123.123.123.123");
user_pref("network.proxy.ftp_port", 8080);
user_pref("network.proxy.http", "123.123.123.123");
user_pref("network.proxy.http_port", 8080);
user_pref("network.proxy.share_proxy_settings", true);
user_pref("network.proxy.ssl", "123.123.123.123");
user_pref("network.proxy.ssl_port", 8080);
user_pref("network.proxy.type", 1);'
    $ProfileFolder = (Get-ChildItem "$env:APPDATA\mozilla\Firefox\Profiles").Name | Where-Object { $_ -like "*Italy" } | Select-Object -First 1
    $ProxyStrings | Out-File -FilePath "$DefaultProfilePath\$ProfileFolder\prefs.js" -Append -Encoding ascii
    Write-Host "Profile setting finished!" -ForegroundColor Green
    Start-Sleep -Seconds 1
}
function Set-Shortcut {
    Write-Host "Setting up shortcut..." -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    $ShortcutLocation = "C:\$env:HOMEPATH\Desktop\Italy.lnk"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutLocation)
    $Shortcut.TargetPath = $FFInstalledPath
    $Shortcut.Arguments = "-P Italy"
    $Shortcut.Save()
    Write-Host "Shortcut set on Desktop" -ForegroundColor Green
    Start-Sleep -Seconds 1
}

if ((Test-Path $DefaultProfilePath) -eq $true) {
    Write-Host "Firefox was used before" -ForegroundColor Green
    Start-Sleep -Seconds 1
    $ItalyProfile = Test-Path -Path "$DefaultProfilePath\*Italy"
    if ($ItalyProfile -eq $true) {
        Write-Host "Profile exists" -ForegroundColor Green
        Start-Sleep -Seconds 1
        Set-Shortcut
    }
    else {
        Write-Host "Profile is missing"
        Start-Sleep -Seconds 1
        Set-Profile
        Set-Shortcut
    }
}
else {
    Write-Host "Firefox wasn't used before, starting now..." -ForegroundColor Yellow
    Start-Sleep -Seconds 1
    Start-Process -FilePath $FFInstalledPath
    Start-Sleep -Seconds 5
    Stop-Process -Name firefox -Force -ErrorAction SilentlyContinue
    Set-Profile
    Set-Shortcut
}