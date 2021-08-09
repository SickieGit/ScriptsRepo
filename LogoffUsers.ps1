<#
.SYNOPSIS
    Logout any disconnected user from the computer

.DESCRIPTION
    The script runs quser command and splits results on each row where 'disc' is mentioned.
    This will create 7 rows per one user with 2nd row being the username
    It will iterate for each user and run loggoff command for every 2nd row of each 7th group.
#>

$sessions = ((quser | Where-Object { $_ -match "disc" }) -split ' +')
$cycle = $sessions.count / 7
$count = 2

For ($i = 0; $i -lt $cycle; $i++) {
    logoff $sessions[$count]
    $count = $count + 7
}