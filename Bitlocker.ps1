<#
.SYNOPSIS
    A script for automation job to get the .csv report of the Bitlocker status on all laptops. Each run iterates on the previous data and updates when/where necessary.

.DESCRIPTION
        LOG CHECK
    Check if \\network log matches backup, if:
    NO > Copy backup and overwrite

        LIST MAINTENANCE
    Get all AD laptops
    Get all .csv laptops
    Check if new exist > Add Name, Owner, AD Bitlocker Policy, BLANK for rest > Save list
    Check if list has not existing ones > export withot them
    Sort list by computer name
    Check if the owners match the AD description > update if not
    Get all AD Bitlocker Policy NO > check if any is YES > Update field > Save list

    Get all TPM_Present thaT are BLANK from list and ping for online.
    Get all online TPM_Present BLANK and try to check remote. If:
        Remote fails > Write comment RDP Issue
        Remote works > Check online if TPM Present, if:
            YES > Write YES to TPM Present
            NO > Write NO to TPM Present, TPM Enabled, Bitlocker Enabled, Fully Encrypted

    Get from list all online Computers where TPM Present = YES > check TPM Enabled, if:
        BLANK or NO > Check online if TPM Enabled, if:
            YES > Write YES to TPM Enabled
            NO > Write NO To TPM Enabled, Bitlocker Enabled, Fully Encrypted

    Get from list all online computers where TPM Enabled = YES > check Bitlocker Enabled, if:
        BLANK or NO > Check online if Bitlocker Enabled is setup, if:
            YES > write YES to Bitlocker Enabled
            NO > write NO to Bitlocker Enabled, Fully Enrypted

    Get from list all online computers where Bitlocker Enabled = YES, if Fully Encrypted:
        BLANK or NO > Check online if Fully Encrypted, if:
            YES > Write YES to Fully Encrypted
            NO > Write NO to Fully Encrypted

    Export to CSV file > Copy new csv file to backup location
    Comments>
    YES-YES-YES-YES-YES = DONE
    YES-YES-YES-YES-NO = Bitlocker enabled, but disk not encrypted fully
    YES-YES-YES-NO-NO = Bitlocker capable, group member, but hasn't been enabled
    YES-YES-NO-NO-NO = Bitlocker capable, group member, TPM and Bitlocker not enabled
    YES-NO-NO-NO-NO = Not capable for Bitlocker, but member of the group, remove
    NO-YES-YES-NO-NO = Bitlocker capable, no group membership, should be setup
    NO-YES-NO-NO-NO = Bitlocker capable, no group membership, TPM and Bitlocker not enabled
    NO-NO-NO-NO-NO = Not capable for Bitlocker, no group membership, leave alone
    ANY-BLANK-BLANK-BLANK-BLANK; comment not BLANK = Always offline, check with owner

    REPORT ALL > If comment = BLANK > Unhandled case, check manually

  .PARAMETER LogNetworkPath
      A location of the main report.

  .PARAMETER LogBackupPath
      A location of the report backup.

  .PARAMETER SearchCriteria
      A search criteria to look up computer hostnames in AD.

  .OUTPUTS
      The script will output a csv file at a location specified in LogNetworkPath variable.
#>

$LogNetworkPath = "\\network\reports\Bitlocker.csv"
$LogBackupPath = "\\network\Bitlocker.csv"

#LOG AND BACKUP CHECK
if ((Test-Path $LogBackupPath) -and (Test-Path $LogNetworkPath)) {
    if (Compare-Object -ReferenceObject $(Get-Content -path $LogNetworkPath) -DifferenceObject $(Get-Content -Path $LogBackupPath)) {
        Copy-Item -Path $LogBackupPath -Destination $LogNetworkPath -Force
    }
}
elseif ((Test-Path $LogBackupPath) -and (!(Test-Path $LogNetworkPath))) {
    Copy-Item -Path $LogBackupPath -Destination $LogNetworkPath -Force
}

#GET ALL AD LAPTOPS AND ALL LIST LAPTOPS
$SearchCriteria = "LAPTOP*"
$AdComputers = Get-ADComputer -filter * | Where-Object { $_.Name -like $SearchCriteria } | Select-Object -ExpandProperty Name
$ListComputers = Import-Csv -Path $LogNetworkPath

#COMPARE AD AND LIST LAPTOPS, ADD NEW LAPTOPS TO THE LIST, REMOVE OLD FROM THE LIST
$Comparison = Compare-Object -ReferenceObject $AdComputers -DifferenceObject $ListComputers.Hostname
if ($Comparison) {
    $AdNewComputers = $Comparison | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object -ExpandProperty InputObject
    $ListDeletedComputers = $Comparison | Where-Object { $_.SideIndicator -eq '=>' } | Select-Object -ExpandProperty InputObject

    foreach ($AdNewComputer in $AdNewComputers) {
        $Owner = Get-ADComputer $AdNewComputer -Properties Description | Select-Object -ExpandProperty Description
        $ComputerGroups = Get-ADComputer $AdNewComputer | Get-ADPrincipalGroupMembership | Select-Object -ExpandProperty Name
        if ($ComputerGroups -contains "AD Bitlocker Policy" ) { $GLBitlocker = "YES" }
        else { $GLBitlocker = "NO" }
        
        $NewRow = "$ADNewComputer,$Owner,$GLBitlocker,BLANK,BLANK,BLANK,BLANK,BLANK"
        $NewRow | Add-Content -Path $LogNetworkPath
    }
    foreach ($ListDeletedComputer in $ListDeletedComputers) {
        $ListComputers | Where-Object { $_.Hostname -ne $ListDeletedComputer } | Export-Csv -Path $LogNetworkPath -Force -NoTypeInformation
    }
}
#SORT BY LAPTOP NAME AND SAVE
Rename-Item -Path $LogNetworkPath -NewName "TempBitlocker.csv"
$TempBitlocker = "\\network\reports\TempBitlocker.csv"
Import-Csv -Path $TempBitlocker  | Sort-Object -Property Hostname | Export-Csv -Path $LogNetworkPath -Force -NoTypeInformation
Remove-Item -Path $TempBitlocker

#FIX DESCRIPTION IF THE OWNER HAS CHANGED
$DescriptionImport = Import-Csv -Path $LogNetworkPath
foreach ($Item in $DescriptionImport) {
    $DescriptionAD = Get-ADComputer $Item.Hostname -Properties Description | Select-Object -ExpandProperty Description
    if ($DescriptionAD -ne $Item.Owner) {
        $RowIndex = [array]::IndexOf($DescriptionImport.Hostname, $Item.Hostname)
        $DescriptionImport[$RowIndex].Owner = "$DescriptionAD"
    }
}
$DescriptionImport | Export-Csv -Path $LogNetworkPath -NoTypeInformation

#IMPORT LIST AND CHECK IF ANY AD Bitlocker Policy MEMBER HAS BEEN JOINED IN THE MEANTIME, THEN EDIT LIST
$NewListComputers = Import-Csv -Path $LogNetworkPath
$ListComputersGlBitlockerNO = $NewListComputers | Where-Object { $_.GL_Bitlocker -eq "NO" } 
ForEach ($ListComputerGlBitlockerNO in $ListComputersGlBitlockerNO) {
    $ComputerSG = Get-ADComputer $ListComputerGlBitlockerNO.Hostname | Get-ADPrincipalGroupMembership | Select-Object -ExpandProperty Name
    if ($ComputerSG -contains "AD Bitlocker Policy" ) {
        $RowIndex = [array]::IndexOf($NewListComputers.Hostname, $ListComputerGlBitlockerNO.Hostname)
        $NewListComputers[$RowIndex].GL_Bitlocker = "YES"
    }
}
$NewListComputers | Export-Csv -Path $LogNetworkPath -Force -NoTypeInformation

#LIST ALL LAPTOPS FROM THE LIST WITH BLANK VALUE FOR TPM_PRESENT
$TPMPresentListBLANK = Import-Csv -Path $LogNetworkPath | Where-Object { $_.TPM_Present -eq "BLANK" }

#PING THEM, KEEP ONLINE ONLY
$PingComputers = $TPMPresentListBLANK | ForEach-Object { 
    Test-Connection -ComputerName $_.Hostname -Count 1 -AsJob } | 
Get-Job | Receive-Job -Wait
$OnlineTPMpresentBLANKlaptops = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | 
Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
Get-Job | Remove-Job

#GET THE UPDATED LIST
$WorkListImport = Import-Csv -Path $LogNetworkPath

#CHECK EACH ONLINE LAPTOP WITH TPM_PRESENT=BLANK IF TPM I PRESENT, EDIT LIST ACCORDINGLY
foreach ($Computer in $OnlineTPMpresentBLANKlaptops) {
    $RowIndex = [array]::IndexOf($WorkListImport.Hostname, $Computer)
    try {
        $TPMpresent = Invoke-Command -ComputerName $computer -ScriptBlock { (Get-TPM).TPMPresent } -ErrorAction Stop
    }
    catch {
        $WorkListImport[$RowIndex].Comment = "RDPIssue"
        Continue
    }
    if ($TPMpresent -eq $True) {
        $WorkListImport[$RowIndex].TPM_Present = "YES"
    }
   
    elseif ($TPMpresent -match $False) {
        $WorkListImport[$RowIndex].TPM_Present = "NO"
        $WorkListImport[$RowIndex].TPM_Enabled = "NO"
        $WorkListImport[$RowIndex].Bitlocker_Enabled = "NO"
        $WorkListImport[$RowIndex].Fully_Encrypted = "NO"
    }
}

#GET ALL LAPTOPS FROM THE LIST WITH TPM_PRESENT = YES and TPM_ENABLED = BLANK or NO
$TPMPresentListYES = $WorkListImport | Where-Object { $_.TPM_Present -eq "YES" -and ($_.TPM_Enabled -eq "BLANK" -or $_.TPM_Enabled -eq "NO") }

#PING THEM, KEEP ONLINE ONLY
$PingComputers = $TPMPresentListYES | ForEach-Object { 
    Test-Connection -ComputerName $_.Hostname -Count 1 -AsJob } | 
Get-Job | Receive-Job -Wait
$OnlineTPMpresentYESlaptops = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | 
Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
Get-Job | Remove-Job

#CHECK EACH ONLINE LAPTOP WITH TPM_PRESENT=BLANK IF TPM IS PRESENT, EDIT LIST ACCORDINGLY
foreach ($Computer in $OnlineTPMpresentYESlaptops) {
    $RowIndex = [array]::IndexOf($WorkListImport.Hostname, $Computer)
    try {
        $TPMenabled = Invoke-Command -ComputerName $computer -ScriptBlock { (Get-TPM).TPMEnabled } -ErrorAction Stop
    }
    catch {
        $WorkListImport[$RowIndex].Comment = "RDPIssue"
        Continue
    }
    if ($TPMenabled -eq $True) {
        $WorkListImport[$RowIndex].TPM_Enabled = "YES"
    }
   
    elseif ($TPMenabled -match $False) {
        $WorkListImport[$RowIndex].TPM_Enabled = "NO"
        $WorkListImport[$RowIndex].Bitlocker_Enabled = "NO"
        $WorkListImport[$RowIndex].Fully_Encrypted = "NO"
    }
}

#GET ALL LAPTOPS FROM THE LIST WITH TPM_ENABLED = YES and Bitlocker_Enabled = BLANK or NO
$BitlockerEnabledYES = $WorkListImport | Where-Object { $_.TPM_Enabled -eq "YES" -and ($_.Bitlocker_Enabled -eq "BLANK" -or $_.Bitlocker_Enabled -eq "NO") }

#PING THEM, KEEP ONLINE ONLY
$PingComputers = $BitlockerEnabledYES | ForEach-Object { 
    Test-Connection -ComputerName $_.Hostname -Count 1 -AsJob } | 
Get-Job | Receive-Job -Wait
$OnlineBitlockerEnabledYES = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | 
Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
Get-Job | Remove-Job

#CHECK EACH ONLINE LAPTOP WITH TPM_ENABLED=BLANK IF TPM IS PRESENT, EDIT LIST ACCORDINGLY
foreach ($Computer in $OnlineBitlockerEnabledYES) {
    $RowIndex = [array]::IndexOf($WorkListImport.Hostname, $Computer)
    try {
        $BitlockerEnabled = Invoke-Command -ComputerName $Computer -ScriptBlock {
            Get-BitLockerVolume -MountPoint C | Select-Object ProtectionStatus
        } | Select-Object -ExpandProperty ProtectionStatus -ErrorAction Stop
    }
    catch {
        $WorkListImport[$RowIndex].Comment = "RDPIssue"
        Continue
    }
    if ($BitlockerEnabled -match "On") {
        $WorkListImport[$RowIndex].Bitlocker_Enabled = "YES"
    }
   
    elseif ($BitlockerEnabled -match "Off") {
        $WorkListImport[$RowIndex].Bitlocker_Enabled = "NO"
        $WorkListImport[$RowIndex].Fully_Encrypted = "NO"
    }
}

#GET ALL LAPTOPS FROM THE LIST WITH TPM_ENABLED = YES and Bitlocker_Enabled = BLANK or NO
$BitlockerEncryptionYES = $WorkListImport | Where-Object { $_.Bitlocker_Enabled -eq "YES" -and ($_.Fully_Encrypted -eq "BLANK" -or $_.Fully_Encrypted -eq "NO") }

#PING THEM, KEEP ONLINE ONLY
$PingComputers = $BitlockerEncryptionYES | ForEach-Object { 
    Test-Connection -ComputerName $_.Hostname -Count 1 -AsJob } | 
Get-Job | Receive-Job -Wait
$OnlineBitlockerEncryptionYES = $PingComputers | Where-Object { $_.StatusCode -eq 0 } | 
Select-Object -ExpandProperty Address | Where-Object { $_ -ne $env:COMPUTERNAME }
Get-Job | Remove-Job

#CHECK EACH ONLINE LAPTOP WITH BITLOCKER_ENCYPTION=YES IF TPM IS PRESENT, EDIT LIST ACCORDINGLY
foreach ($Computer in $OnlineBitlockerEncryptionYES) {
    $RowIndex = [array]::IndexOf($WorkListImport.Hostname, $Computer)
    try {
        $BitlockerEncryption = Invoke-Command -ComputerName $Computer -ScriptBlock {
            Get-BitLockerVolume -MountPoint C | Select-Object EncryptionPercentage
        } | Select-Object -ExpandProperty EncryptionPercentage -ErrorAction Stop
    }
    catch {
        $WorkListImport[$RowIndex].Comment = "RDPIssue"
        Continue
    }
    if ($BitlockerEncryption -eq 100) {
        $WorkListImport[$RowIndex].Fully_Encrypted = "YES"
    }
   
    elseif ($BitlockerEncryption -lt 100) {
        $WorkListImport[$RowIndex].Fully_Encrypted = "NO"
    }
}

#FINAL EXPORT CSV
$WorkListImport | Export-Csv -Path $LogNetworkPath -NoTypeInformation

#REPORT IMPORT AND GENERATING 
$ReportImport = Import-Csv -Path $LogNetworkPath

foreach ($Entry in $ReportImport) {
    $RowIndex = [array]::IndexOf($ReportImport.Hostname, $Entry.Hostname)
    if ($Entry.GL_Bitlocker -eq "YES" -and 
        $Entry.TPM_Present -eq "YES" -and 
        $Entry.TPM_Enabled -eq "YES" -and 
        $Entry.Bitlocker_Enabled -eq "YES" -and 
        $Entry.Fully_Encrypted -eq "YES") {
        $ReportImport[$RowIndex].Comment = "DONE"
    }
    elseif ($Entry.GL_Bitlocker -eq "YES" -and 
        $Entry.TPM_Present -eq "YES" -and 
        $Entry.TPM_Enabled -eq "YES" -and 
        $Entry.Bitlocker_Enabled -eq "YES" -and 
        $Entry.Fully_Encrypted -eq "NO") {
        $ReportImport[$RowIndex].Comment = "Disk not 100% encrypted"
    }
    elseif ($Entry.GL_Bitlocker -eq "YES" -and 
        $Entry.TPM_Present -eq "YES" -and 
        $Entry.TPM_Enabled -eq "YES" -and 
        $Entry.Bitlocker_Enabled -eq "NO" -and 
        $Entry.Fully_Encrypted -eq "NO") {
        $ReportImport[$RowIndex].Comment = "Set BL locally"
    }
    elseif ($Entry.GL_Bitlocker -eq "YES" -and 
        $Entry.TPM_Present -eq "YES" -and 
        $Entry.TPM_Enabled -eq "NO" -and 
        $Entry.Bitlocker_Enabled -eq "NO" -and 
        $Entry.Fully_Encrypted -eq "NO") {
        $ReportImport[$RowIndex].Comment = "Enable TPM and BL"
    }
    elseif ($Entry.GL_Bitlocker -eq "YES" -and 
        $Entry.TPM_Present -eq "NO" -and 
        $Entry.TPM_Enabled -eq "NO" -and 
        $Entry.Bitlocker_Enabled -eq "NO" -and 
        $Entry.Fully_Encrypted -eq "NO") {
        $ReportImport[$RowIndex].Comment = "Remove from AD Bitlocker Policy"
    }
    elseif ($Entry.GL_Bitlocker -eq "NO" -and 
        $Entry.TPM_Present -eq "YES" -and 
        $Entry.TPM_Enabled -eq "YES" -and 
        $Entry.Bitlocker_Enabled -eq "NO" -and 
        $Entry.Fully_Encrypted -eq "NO") {
        $ReportImport[$RowIndex].Comment = "Add to AD Bitlocker Policy, set BL"
    }
    elseif ($Entry.GL_Bitlocker -eq "NO" -and 
        $Entry.TPM_Present -eq "YES" -and 
        $Entry.TPM_Enabled -eq "NO" -and 
        $Entry.Bitlocker_Enabled -eq "NO" -and 
        $Entry.Fully_Encrypted -eq "NO") {
        $ReportImport[$RowIndex].Comment = "Add to AD Bitlocker Policy, enable TPM, set BL"
    }
    elseif ($Entry.GL_Bitlocker -eq "NO" -and 
        $Entry.TPM_Present -eq "NO" -and 
        $Entry.TPM_Enabled -eq "NO" -and 
        $Entry.Bitlocker_Enabled -eq "NO" -and 
        $Entry.Fully_Encrypted -eq "NO") {
        $ReportImport[$RowIndex].Comment = "IGNORE"
    }
    elseif ($Entry.Comment -eq "BLANK" -or
        $Entry.Commment -eq "Offline" -and 
        $Entry.TPM_Present -eq "BLANK" -and 
        $Entry.TPM_Enabled -eq "BLANK" -and 
        $Entry.Bitlocker_Enabled -eq "BLANK" -and 
        $Entry.Fully_Encrypted -eq "BLANK") {
        $ReportImport[$RowIndex].Comment = "Offline"
    }
    elseif ($Entry.Comment -eq "BLANK" -and
        $Entry.TPM_Present -ne "BLANK" -or 
        $Entry.TPM_Enabled -ne "BLANK" -or 
        $Entry.Bitlocker_Enabled -ne "BLANK" -or 
        $Entry.Fully_Encrypted -ne "BLANK") {
        $ReportImport[$RowIndex].Comment = "UNHANDLED CASE"
    }
}
$ReportImport | Export-Csv -Path $LogNetworkPath -NoTypeInformation

#BACKUP COPYING
Copy-Item -Path $LogNetworkPath -Destination $LogBackupPath -Force | Out-Null