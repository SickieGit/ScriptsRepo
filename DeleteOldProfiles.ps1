<#
.SYNOPSIS
    The script job will delete any user folde with its matching regkey that is no longer present in AD.

.DESCRIPTION
    The script will get the list of specified computers and for each one get the list of all user profiles
    If the user profile name isn't enabled in AD, it will be deleted alongside its matching registry key.

.OUTPUTS
    The script will output a the current operation to the console window.
#>

$computers = Get-Content "\\network\List.txt"
foreach ($computer in $computers) {
    $LocalAccounts = (Get-ChildItem "\\$computer\c$\Users\").Name | Where-Object { $_ -like "*.*" }
    foreach ($LocalAccount in $LocalAccounts) {
        try {
            $ADAccount = (Get-ADUser -Identity $LocalAccount).Enabled
            $UserExists = $True
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $UserExists = $False
        }
        if ($UserExists -eq $True -and $ADAccount -eq $True) {
            Write-Host "$LocalAccount is active" 
        }
        else {
            $session = New-PSSession -ComputerName $computer
            Invoke-Command -Session $session -ScriptBlock { 
                Write-Host "Deleting folder $using:LocalAccount..."
                $LocalAccountPath = "C:\users\$using:LocalAccount"
                cmd /c rmdir /s /q $LocalAccountPath

                Write-Host "Deleting regkey $using:LocalAccount..."
                $path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
                $PSChildNames = (Get-Item $path).GetSubKeyNames()
                foreach ($PSChildName in $PSChildNames) {
                    $FullRegPath = $path + "\" + $PSChildName
                    $RegSearch = Get-ItemProperty -Path $FullRegPath
                    if ($RegSearch.ProfileImagePath -eq $LocalAccountPath) {
                        Remove-Item -Path $FullRegPath -Recurse -Force
                    }
                }
            }
            Get-PSSession | Remove-PSSession
            Write-Host "Done!" -ForegroundColor Green
        }
    }
}