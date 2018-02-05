# =====================================================================================
# .NAME == OfficeActivationScript.ps1
# .DESCRIPTION == PowerShell script to manually activate Office 2016 to a KMS.
#                 It will also tear down the activation panel when opening Office produce due to useless OEM keys.
# .INSTRUCTIONS == 1) Run: PowerShell x64 as Admin
#                  2) Ensure: Set-ExecutionPolicy RemoteSigned
#                  3) Run: c:\path\to\OfficeActivationScript.ps1
# .AUTHOR == Adrienne Love
# .DATE == June 2017
# .VERSION == v0.5 Alpha
# .IMPROVEMENTS == a) Variable the office folder path and registry path to support more than Office 2016
#                  b) Make it portable to different KMS FQDN
#                  c) Add test for checking if Office processes are running, and if they are
#                     give the user an opportunity to save open work before continuing
#                     Then, test for running Office processes again, in case the user also closed them themselves
#                     If Office processes are still running, close them
#                     Then allow the user to press enter to continue the script
#                     This set of events should run at the very beginning of this script     
# =====================================================================================

# =====================================================================================
# INITIALIZE AND DECLARE VARIABLES
$kmsHostName = ""
$officeDirectoryPath = "c:\program files (x86)\microsoft office\office16" 
$registryKeys = "hklm\software\microsoft\office\16.0\common\oem","hklm\software\wow6432node\microsoft\office\16.0\common\oem"
$registryExportPath = "c:\windows\temp"
$registryBackupFile = "OfficeOEMKeyBackup.reg"
#$processes = "WINWORD","EXCEL","POWERPNT","MSACCESS","ONENOTE","ONENOTEM","OUTLOOK","MSPUB","LYNC"  HARDCODING IS BAD FOR YOU
$processes = Get-Process | Where {$_.Path -like "*Microsoft Office*"} | Select ProcessName -ExpandProperty ProcessName
# =====================================================================================

# =====================================================================================
# BEGIN RESUABLE FUNCTIONS

# Test if KMS is reachable
function testKMSConnection ($kmsHostName){
    Write-Host "`nTesting connection to $kmsHostname ..."
    if (!(Test-Connection $kmsHostName -Quiet)){
        Write-Host "Could not reach $kmsHostname.  Are you on the Network?"
        return $false
        Exit
        }
    else { 
        Write-Host "$kmsHostname exists." 
        return $true
        }
    }

# Test if the Office location exists
function verifyOfficeLocation ($officeDirectoryPath){
    Write-Host "`nTesting Office exists..."
    if (Test-Path $officeDirectoryPath -ErrorAction "SilentlyContinue"){
        Set-Location $officeDirectoryPath
        Write-Host "Office found.  Set path to $officeDirectoryPath`n"
        return $true
        }
    else { 
        Write-Host "$officeDirectoryPath does not exist."
        return $false
        Exit 
        }
    }

# Test if Registry Keys exist
function verifyRegistryKeys ($registryKey) {
    Write-Host "`nTesting Registry Keys exist..."
    if (Test-Path -Path "Registry::$registryKey") {
        Write-Host "$registryKey exists."
        return $true
        }
    else { 
        Write-Host "$registryKey does not exist."
        return $false 
        }
}

# END REUSABLE FUNCTIONS
# =====================================================================================

Clear-Host
Write-Host "Tip: CTRL+C will abandon this script at any time."

# =====================================================================================
# BEGIN USER INPUT CYCLE
while (($kmsHostName -ne "redacted") -and ($kmsHostname -ne "www.google.com")) { #TODO: Don't hardcode that shit.  Make it easy for someone to choose 1 of 3, but hide it from client
    $kmsHostName = Read-Host "The KMS server path"
    }
# END USER INPUT CYCLE
# =====================================================================================

# =====================================================================================
# BEGIN SCRIPT CYCLE

Write-Host "`nBefore we begin, I have to check for running Office programs and close them."
foreach ($process in $processes) {
    Write-Warning "`n$process is open."
    $continue = Read-Host "Please save your work, then press Enter here so I can close it and continue"
    Stop-Process -Name $process  
    }

testKMSConnection $kmsHostName
verifyOfficeLocation $officeDirectoryPath

& cscript ospp.vbs /sethst:$kmsHostName
& cscript ospp.vbs /act
Start-Sleep -Seconds 5
& cscript ospp.vbs /dstatus

foreach ($registryKey in $registryKeys) {
    $state = verifyRegistryKeys $registryKey
    if ($state -eq $true) {
        Write-Host "`nExporting $registryKey for backup purposes..."
        $i++
        & reg export $registryKey $registryExportPath\$i.reg /y
        Remove-Item "Registry::$registryKey"
        Write-Host "`nRemoved the registry key: $registryKey"
    }
}


Get-Content "$registryExportPath\*.reg" | Set-Content "$registryExportPath\$registryBackupFile"
Remove-Item "$registryExportPath\*.reg" -Exclude "$registryBackupFile"
np $registryExportPath\$registryBackupFile


Write-Host "`nFiring up Word. See if you get an activation panel and also check File -> Account that product is Activated."
Start-Process -FilePath "c:\program files (x86)\microsoft office\root\office16\winword.exe"






# END SCRIPT CYCLE
# =====================================================================================



# FOR DEBUGGING BECAUSE I'M LAZY; RESET WHAT I HAD BEFORE TESTING THE SCRIPT
foreach ($registryKey in $registryKeys) {
    New-Item -Path "Registry::$registryKey"
    }
Remove-Item "$registryExportPath\*.reg"
Set-Location "c:\Scripts\PowerShell"

