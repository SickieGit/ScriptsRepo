<#
  .SYNOPSIS
      Generate a random password from charsets defined.

  .DESCRIPTION
      The script will get the defined number of small caps letters, all caps and numbers,
      and then scramble them randomly to get the random password.
      At the end, it will automatically loaded into the clipboard for ease of use.

  .OUTPUTS
      The script outputs to the console window the random password and loads it into clipboard.
#>

function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs = ""
    return [String]$characters[$random]
}

function Edit-StringOrder([string]$inputString) {     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

$password = Get-RandomCharacters -length 5 -characters 'abcdefghiklmnoprstuvwxyz'
$password += Get-RandomCharacters -length 3 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$password += Get-RandomCharacters -length 2 -characters '1234567890'

$password = Edit-StringOrder $password

Write-Host "$password" -ForegroundColor Green
Set-Clipboard $password