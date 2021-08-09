function  Get-FolderAcl {
  <#
  .SYNOPSIS
      Get the list of users/groups and their rights over a particular folder.

  .DESCRIPTION
      The script will get the list of all users or groups that have allowed or explicitly blocked access over a particular folder,
      then it will exclude the names of the users that are specificed,
      then it will recursively go 1 or more levels deeper into the child folders and get those users/groups
      then it will output the report as a csv file at the location specified. 

  .PARAMETER RootPath
      Enter the root folder path for scanning. Mandatory.

  .PARAMETER OutFile
      Enter the output destination of the csv file alongisde name. Default is user's desktop.

  .PARAMETER RemovedFromArray
      Enter one or more names for the users that should be ommited from the report. Default is empty.

  .PARAMETER Level
      Enter how many levels recursevely should scan go deep. 0 means 1 level, 1 means 2 levels, etc. Default is 0.

  .OUTPUTS
      The script will output a csv file at a location specified in OutFile parameter.

  .EXAMPLE
      Get-FolderAcl -RootPath C:\users

      Description
      Will retrieve a list of all users that have granted or explicitly blocked rights over the C:\users folder,
      it won't omit any usernames, it will go recursevely 1 level down and scan the subfolders with C:\users,
      and it will output it to your desktop as GetFolderAcl.csv
  
  .EXAMPLE
      Get-FolderAcl -RootPath C:\users -OutFile C:\temp\export.csv -RemovedFromArray "NT AUTHORITY\SYSTEM","NT AUTHORITY\BATCH","NT AUTHORITY\INTERACTIVE","NT AUTHORITY\SERVICE" -Level 1

      Description
      Will retrieve a list of all users that have granted or explicitly blocked rights over the C:\users folder,
      will omit the NT AUTHORITY listed users, will go 2 levels down within C:\users folder
  #>
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
    [string]$RootPath,
    [string]$OutFile = "$env:HOMEDRIVE$env:HOMEPATH\Desktop\GetFolderAcl.csv",
    [string[]]$RemovedFromArray,
    [int]$Level = 0
  )

  $Folders = Get-ChildItem $RootPath -Depth $Level | Where-Object { $_.psiscontainer -eq $true }
  $Folders | ForEach-Object {
    $fpath = $_.FullName
    Get-Acl $fpath | Select-Object -Expand Access |
    Select-Object @{
      n = 'FolderName'; e = { $fpath }
    },
    IdentityReference,
    AccessControlType,
    FileSystemRights | where-object {
      $RemovedFromArray -notcontains $_.IdentityReference
    }
  } | Export-Csv $OutFile -NoTypeInformation
}