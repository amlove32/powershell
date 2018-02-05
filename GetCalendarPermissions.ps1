# ==================================================================================================================================
# ==================================================================================================================================
# .NAME == GetCalendarPermissions.ps1
#
# .DESCRIPTION ==  An Exchange PowerShell Script that will find all the calendars that a user has permission to.
#                  It asks for the script runner's input for which username to search the Exchange db for.
#                  It then searches calendar permissions of all users, looking for the user from the input.
#
# .INSTRUCTIONS == 1)  Run: Exchange PowerShell x64 as Admin
#                  2)  Ensure: Set-ExecutionPolicy RemoteSigned
#                  3)  Run:  cd c:\path\to\script
#                  4)  Run:  .\GetCalendarPermissions.ps1
#
# .AUTHOR == Adrienne Love
#
#. DATE == July 2017
#
# .VERSION == v0.5 Alpha
#
# .IMPROVEMENTS == (a) ERROR HANDLING TODO
#                  (b) Search in batches TODO
#                  (c) Make it more robust with searching other folders than calendar (prompt the runner) TODO
#
# ==================================================================================================================================
# ==================================================================================================================================


# =================================================================
# FUNCTIONS


# Test for valid user input.
function TestUserExists ($user){
    Write-Host "`nChecking $user exists..."
    if (Get-Mailbox -Identity $user -ErrorAction "SilentlyContinue){
        Write-Host "`nFound $user... valid entry...`n"
        return $true
        }
    else {
        Write-Host "`n$user does not exist, try again...`n"
        return $false
        }
    }


# =================================================================
# BEGIN USER INPUT CYCLE

$user = ""

Write-Host "ProTip:  CTRL+C will exit the script at any stage.`nHi, $env:username `nI will find all the calendars that a user has permission to access."

# This loop is to force the runner to give me valid input; else no reason to run the script.
do {
    while (($user = Read-Host "`nEnter the username you want to search for:") -eq "") {}
    $exist = TestUserExists $user
    }
until ($exist -eq $true)


# =================================================================
# FIRE UP THE SCRIPT

# Why export the found boxes to a txt file and import it later?  Because having an array >1000 items in a var is a terrible idea.

# Put the files on the runner's desktop because we might not be able to save to root
$saveFilePath = "c:\users\$env:username\desktop"

# This is here because stats are fun.
$stopwatch = [system.diagnostics.stopwatch]::startNew()
$i = 0
$count = 0
$skip = 0
# $numberOfBoxes = ((Get-Mailbox -ResultSize Unlimited).count)

Write-Host "`nI'm gathering all the mailboxes on Exchange in sets of 500, and checking those.  This might take several minutes.`n"

# The while loop is batch control.  We get 500 objects at a time.  If 500+1 is null, exit the loop
# I need to rethink this TODO
while ((Get-Mailbox | select $_ -Skip $skip -First 1) -ne $null) {
    Get-Mailbox  | select $_ -ExpandProperty Alias -Skip $skip -First 500 | %{($_ + ":\Calendar")} | Out-File $saveFilePath\ExchangeData.txt -Force
    Get-Content $saveFilePath\ExchangeData.txt | 
    foreach {
        $i++
        if (Get-MailboxFolderPermission -identity $_ -user $user -ErrorAction "SilentlyContinue") {
            $count++
            $_ | Out-File $saveFilePath\FoundCalendars.txt -Force -Append
            Get-MailboxFolderPermission -identity $_ -user $user | Select User,AccessRights | Out-File $saveFilePath\FoundCalendars.txt -Force -Append
        }
    $skip = $skip + 500
    }
}
 
$stopwatch.Stop()
Write-Host "`nI searched $i calendars for $user.`nI found $user on $count calendars.`nIt took $stopwatch.Elapsed.TotalSeconds seconds to run this script.`n`nNow clean-up..."
Remove-Item $saveFilePath\ExchangeData.txt
Write-Host "`nOpening your data file...`n"
Invoke-Item $saveFilePath\FoundCalendars.txt
