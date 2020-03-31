<#
.SYNOPSIS

  Installs the DQA ESM Program

.DESCRIPTION

  Installs the DQA ESM Program

  1. Creates "$env:userprofile\AppData\Roaming\DQAESM" Directory
  2. Downloads DQA ESM Package from Confgured URL
  3. Unpacks DQA ESM Package to current user appdata directory
  4. Creates an hourly scheduled tasks running as current user

.NOTES

  Version:        1.0.0.0
  Author:         Jean-Pierre Simonis - Delivery Quality Assurance (DQA)
  Creation Date:  20200309
  Purpose/Change: Initial Release

.EXAMPLE

  .\Deploy-DQAESM.ps1


#>

###########################
#    ESM Configuration    #
###########################

#Email Signature Block Scheduled Task Names
$ScheduledESMTaskName = "DQAESM-Updater"
$ScheduledESMTaskDescription = "DQA Email Signature Manager (App & Configuration)"

#Email Signature Block Scheduled Task Names
$ScheduledESMSignatureTaskName = "DQAESM-Signature-Updater"
$ScheduledESMSignatureTaskDescription = "DQA Email Signature Manager (Signature Block Template Management)"

#Scheduler Interval (minutes)
$ScheduledInterval = 60

#Installation Path
$destinationFolder = "$env:userprofile\AppData\Roaming\DQAESM"

#Download Locations
$DQAESMPackageName = "DQAESM"

$DQAESMPackage = "https://sample.com/esm/$DQAESMPackageName.zip"
$DQAESMPackageMD5 = "https://sample.com/esm/$DQAESMPackageName.md5"


###########################
#       Functions         #
###########################

#Function to extract zip files
function unzip ($file,$ExtractPath) {
    expand-archive -Path $file -DestinationPath $ExtractPath -Force
 }


###########################
#    Script Execution     #
###########################


# Create Program Files Folder
if (-not(test-path $destinationFolder)) {
    try {
        $null = New-Item -type "Directory" -Path $destinationFolder -Force
        $null = New-Item -type "Directory" -Path $destinationFolder\Templates -Force

    }
    catch {
        Write-Output "Failed to create program folder $destinationFolder"
        Exit 911
    }
}

# Download DQA ESM Script Package

   Invoke-WebRequest -Method GET -Uri $DQAESMPackageMD5 -OutFile "$destinationFolder\$DQAESMPackageName.md5" -UseBasicParsing
   Invoke-WebRequest -Method GET -Uri $DQAESMPackage -OutFile "$destinationFolder\$DQAESMPackageName.zip" -UseBasicParsing

# Unzip DQA ESM Script Package

try { $ExtractScriptPack = Unzip "$destinationFolder\$DQAESMPackageName.zip" $destinationFolder }
catch {
    Write-Output "Failed to extract the DQA ESM Script Pack to $destinationFolder"
    Exit 912
}

# Remove Existing Scheduled Task if Present
$scheduledTask = Get-ScheduledTask -TaskName $ScheduledESMTaskName -erroraction silentlycontinue
If ($scheduledTask) {
    try { Unregister-ScheduledTask -taskname $ScheduledESMTaskName -confirm:$false }
    catch {
        Write-Output "Failed to remove the existing Scheduled Task named $ScheduledESMTaskName."
        Exit 914
    }
}
$scheduledTask = Get-ScheduledTask -TaskName $ScheduledESMSignatureTaskName -erroraction silentlycontinue
If ($scheduledTask) {
    try { Unregister-ScheduledTask -taskname $ScheduledESMSignatureTaskName -confirm:$false }
    catch {
        Write-Output "Failed to remove the existing Scheduled Task named $ScheduledESMSignatureTaskName."
        Exit 914
    }
}

# Register the New Scheduled Tasks
try {
    Write-Output "Installing DQA Email Signature Updater Scheduled Tasks."
    $scriptpath = "$destinationFolder\run-dqaesmupdate.vbs"
    $scriptWorkingDir = "$destinationFolder"
    $trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes $ScheduledInterval) -once -At (Get-Date) -RandomDelay (New-TimeSpan -Minutes 2)
    $action = New-ScheduledTaskAction -Execute 'wscript.exe' -WorkingDirectory $scriptWorkingDir -Argument "//nologo `"$scriptpath`""
    $settings = New-ScheduledTaskSettingsSet -RunOnlyIfNetworkAvailable -AllowStartIfOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -Hidden -Compatibility Win8
    $null = Register-ScheduledTask -taskname $ScheduledESMTaskName -Action $action -Settings $settings -Trigger $trigger -Description $ScheduledESMTaskDescription
    $scriptpath = "$destinationFolder\run-dqaesmsigupdate.vbs"
    $action = New-ScheduledTaskAction -Execute 'wscript.exe' -WorkingDirectory $scriptWorkingDir -Argument "//nologo `"$scriptpath`""
    $null = Register-ScheduledTask -taskname $ScheduledESMSignatureTaskName -Action $action -Settings $settings -Trigger $trigger -Description $ScheduledESMSignatureTaskDescription

}
catch {
    Write-Output "Failed to register the DQA ESM Scheduled Task."
    Exit 915
}


# Completion
Write-Output "DQA ESM client installation complete."

# Execute the Scheduled Tasks
Start-ScheduledTask -TaskName $ScheduledESMSignatureTaskName


