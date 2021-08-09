function Copy-Client {
    <#
  .SYNOPSIS
      Copy the CCM intallation files to the target computer

  .DESCRIPTION
      The script will copy the prepared installation folder to the C: drive of the target computer

  .PARAMETER Computer
      The name of the remote computer to copy the files to.

  .OUTPUTS
      The script will output a the information on the progress and outcome of the copying process.

  .EXAMPLE
      Copy-Client -Computer TestPC

      Description
      Will copy the CCM installation files to the TestPC's C:\Client folder
  #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$Computer
    )
    $FromDestination = "\\network\SCCM\Client"
    $ToDestination = "\\$Computer\c$\Client"
    Write-Host "Copying Client folder to $Computer..."
    Copy-Item $FromDestination $ToDestination -Recurse -Force
    $TotalFilesFrom = "(Get-ChildItem -Path $FromDestination -Recurse | Measure-Object).Count"
    $TotalFIlesTo = "(Get-ChildItem -Path $ToDestination -Recurse | Measure-Object).Count"
    if ($TotalFilesFrom -eq $TotalFIlesTo) {
        Write-Host "Copying finished successfully"
    }
    else {
        Write-Host "Copying failed"
        Exit
    }
}
function Install-CCM {
    <#
  .SYNOPSIS
      Install the ccm software on a target remote computer.

  .DESCRIPTION
      The script will install the CCM software on a target remote computer from the previously copied installation folder
      The instaltion is silent and the script will just make sure that it has actually started

  .PARAMETER Computer
      The name of the remote computer to install the CCM software to.

  .OUTPUTS
      The script will output a the information on the start of the installation process.

  .EXAMPLE
      Install-CCM -Computer TestPC

      Description
      Will install the CCM software  to the TestPC.
  #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$Computer
    )
    Write-Host "Initiating CCM client installation..."
    Invoke-Command -ComputerName $Computer -ScriptBlock {
        $CCM_Install = "ccmsetup.exe /mp:ccmserver.domain /logon SMSSITECODE=COMPANY"
        & cmd /c $CCM_Install
    }
    for ($i = 1; $i -le 3; $i++) {
        Start-Sleep -Seconds 3
        $ccmcacheExists = Test-Path -Path \\$computer\c$\Windows\ccmcache
        if ($ccmcacheExists -eq $True) {
            Write-Host "The installation has begun, moving on..."
            Break
        }
    }
    if (($i -eq 4) -and ($ccmcacheExists -eq $False)) {
        Write-Host "The installation hasn't started, try again later"
    }
}
function Uninstall-CCM {
    <#
  .SYNOPSIS
      Uninstall the ccm software on a target remote computer.

  .DESCRIPTION
      The script will uninstall the CCM software on a target remote computer.
      The uninstaltion is silent and the script will just make sure that it has actually started

  .PARAMETER Computer
      The name of the remote computer to uninstall the CCM software from.

  .OUTPUTS
      The script will output a the information on the start of the uninstallation process.

  .EXAMPLE
      Uninstall-CCM -Computer TestPC

      Description
      Will uninstall the CCM software from the TestPC.
  #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$Computer
    )
    Write-Host "Uninstalling CCM client..."        
    Invoke-Command -ComputerName $Computer -ScriptBlock {
        $CCM_Uninstall = "c:/Windows/ccmsetup/ccmsetup.exe /uninstall"
        & cmd /c $CCM_Uninstall
    }
    Write-Host "The process will continue in the background"
}