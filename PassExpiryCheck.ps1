function Get-ADPassInfo {
    <#
  .SYNOPSIS
      The function will get the information about the user's AD account pass set date, expiry dat and lock status.

  .DESCRIPTION
      The function first checks if the account exists. If yes, checks for dates when the pass was last set, and when it expires.
      It will then check if the account is locked and offer to unlock it in case it is.

  .PARAMETER Username
      Enter the username for the user that you need to check. Default username is of the user that runs the script.

  .OUTPUTS
      The script will output a console window information about the password set and expiry date, with lock status.

  .EXAMPLE
      Get-ADPassInfo -Username TestUser1

      Description
      Will check the TestUser1 password set and expiry date, color code them for easier readout and check if the account is locked.
      In case it is, will offer to unlock it or not.
  #>
    [CmdletBinding()]
    param (
        [ValidateNotNullOrEmpty()]
        [string]$Username = $env:UserName
    )
    try {
        $query = Get-ADUser $username -Properties "PasswordExpired", "LockedOut", "PassWordLastSet"
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "Account doesn't exist" -ForegroundColor Red
        Break
    }
    $DateFormat = "dddd, dd-MMM-yyyy HH:mm:ss"
    $PasswordLastSet = $query.PasswordLastSet.ToString("$DateFormat")
    $PasswordExpires = $query.PasswordLastSet.AddDays(60).ToString("$DateFormat")

    if ($query.PasswordExpired -eq $True) {
        Write-Host "`nPassword last set: $PasswordLastSet"
        Write-Host "Password expired: $PasswordExpires" -ForegroundColor Red
    }
    else {
        Write-Host "`nPassword last set: $PasswordLastSet"
        Write-Host "Password expires: $PasswordExpires" -ForegroundColor Green
    }

    if ($query.LockedOut -eq $True) {
        Write-Host "Account locked" -ForegroundColor Red
        while (($question = Read-Host "Unlock account? [1]-Yes [2]-No ") -notmatch '^1$|^2$') {
            write-host "Incorrect entry, try again." -ForegroundColor Red 
        }
        if ($question -eq "1") {
            Unlock-ADAccount $username
            Write-Host "Account unlocked" -ForegroundColor Green
        }
        if ($question -eq "2") {
            Write-Host "Account remains locked" -ForegroundColor Red
        }
    }
    else {
        Write-Host "`nThe account isn't locked" -ForegroundColor Green
    }
    Remove-Variable * -ErrorAction SilentlyContinue
    Remove-Module ActiveDirectory
}