# ==================================================================================================================================
# ==================================================================================================================================
# .NAME == ExchangeCalendarPermissions.ps1
#
# .DESCRIPTION ==  A PowerShell Script that will Add, Edit, or Remove a user's access to another user's Exchange calendar.
#                  It gives the runner an interface to specify whose calendar will be edited, who needs permissions, and which access level.
#                  It tests the objects exist, runs the appropriate Exchange cmdlet based on the input, and displays results.
#                  It also logs everything to a file in c:\
#
# .INSTRUCTIONS == 1)  Run: Exchange PowerShell x64 as Admin
#                  2)  Ensure: Set-ExecutionPolicy RemoteSigned
#                  3)  Run: c:/path/to/script/ExchangeCalendarPermissions.ps1
#                  4)  Follow prompts
#
# .AUTHOR == Adrienne Love
#
#. DATE == June 2017
#
# .VERSION == v0.5 Alpha
#
# .IMPROVEMENTS == (a) Refactor it so that more than one user can be added to a calendar at once, for same permission level
#                  (b) ERROR HANDLING
# ==================================================================================================================================
# ==================================================================================================================================

# =================================================================
# BEGIN REUSABLES

# Test mailbox calendar exists
function TestUserMailboxExist ($user) {
    $calendarPath = ($user + ":\Calendar")
    Write-Host "`nChecking $user exists..."
    if (!(Get-MailboxCalendarFolder -Identity $calendarPath -ErrorAction "SilentlyContinue")){
        Write-Warning "Did not find $user. Maybe $user doesn't exist, or something happened to the calendar."
        return $false
        }
    else {
        Write-Host "$user exists."
        return $true
        }
    }

# Test user exists on calendar permissions
function TestPermissionsExist ($userCalendar,$userPermissions) {
    $calendarPath = ($userCalendar + ":\Calendar")
    Write-Host "`nChecking $userPermissions exists on $calendarPath..."
    if (Get-MailboxFolderPermission -Identity "$calendarPath" -User "$userPermissions" -ErrorAction "SilentlyContinue"){
        Write-Host "$userPermissions exists on $calendarPath"
        return $true
        }
    else { 
        Write-Host "$userPermissions does not exist on $calendarPath"
        return $false 
        }
    }

# Get current user permission state on a calendar.  While Set- continues without a terminating error, the legit way to do this is to get the current state and compare it to the desired state then decide what to do
function GetCurrentPermissionState ($userCalendar,$userPermissions) {
    $calendarPath = ($userCalendar + ":\Calendar")
    $currentLevel = Get-MailboxFolderPermission -Identity "$calendarPath" -User "$userPermissions" | Select $_ -ExpandProperty AccessRights
    return $currentLevel
    }

# END REUSABLES
# =================================================================

# =================================================================
# SCRIPT START
# Initialize user input variables
$userNameOfCalendar = ""
$userNameToAssignPerms = ""
$permissionLevel = ""
$continue = ""

Clear-Host

# =================================================================
# BEGIN USER INPUT CYCLE

do {
    while (($userNameOfCalendar = Read-Host "`nEnter the calendar's username") -eq "") {}
    $exist = TestUserMailboxExist $userNameOfCalendar
    }
until ($exist -eq $true)

do {
    while (($usernameToAssignPerms = Read-Host "`nEnter the username who needs access") -eq ""){}
    $exist = TestUserMailboxExist $usernameToAssignPerms
    }
until ($exist -eq $true)

while (($permissionLevel -ne "editor") -and ($permissionLevel -ne "owner") -and ($permissionLevel -ne "reviewer") -and ($permissionLevel -ne "none")){
    $permissionLevel = Read-Host "`nEnter the permission level the user should have (owner, editor, reviewer, none)"
    }

while (($continue -ne "y") -and ($continue -ne "n")){
    $continue = Read-Host "`nGive $usernameToAssignPerms $permissionLevel access to the calendar of $userNameOfCalendar ? (y/n)" # English words
    }
if ($continue -eq "n"){
    Exit
    }

# END USER INPUT CYCLE
# =================================================================

# Generate User Mailbox Calendar Path Var
$userCalendarPath = ($userNameOfCalendar + ":\Calendar")

# First check if the user objects exist
$state = TestPermissionsExist $userNameOfCalendar $userNameToAssignPerms

# =================================================================
# BEGIN ACTION CYCLE
if ($permissionLevel -ne "none") {
    if ($state -eq $false){
        Write-Host "`nAdding $userNameToAssignPerms to $userCalendarPath..."
        Add-MailboxFolderPermission -Identity "$userCalendarPath" -User "$userNameToAssignPerms" -AccessRights "$permissionLevel" -ErrorAction "SilentlyContinue"
        }
    else {
        Write-Host "`nChanging $userNameToAssignPerms access to $userCalendarPath to $permissionLevel..."
        if ((GetCurrentPermissionState $userNameOfCalendar $userNameToAssignPerms) -eq $permissionLevel){
            Write-Warning "$userNameToAssignPerms already has $permissionLevel access to $userCalendarPath"
            Exit
            }
        else {
        Set-MailboxFolderPermission -Identity "$userCalendarPath" -User "$userNameToAssignPerms" -AccessRights "$permissionLevel" -ErrorAction "SilentlyContinue"
        Get-MailboxFolderPermission -Identity "$userCalendarPath" -User "$userNameToAssignPerms"
        }
        }
    }
elseif ($permissionLevel -eq "none"){  # using elseif rather than else here, for debug purposes of testing phase
    if ($state -eq $false) {
        Write-Warning "I can't remove $userNameToAssignPerms from $userCalendarPath.  The user already doesn't exist!"
        }
    else {
        Write-Host "`nRemoving $userNameToAssignPerms from $userCalendarPath..."
        Remove-MailboxFolderPermission -Identity "$userCalendarPath" -User "$userNameToAssignPerms" -ErrorAction "SilentlyContinue" -Confirm:$false
        Write-Host "$userNameToAssignPerms was removed from $userCalendarPath..."
        Get-MailboxFolderPermission -Identity "$userCalendarPath"
        }
    }
# END ACTION CYCLE
# =================================================================
