<#
.SYNOPSIS

  Downloads and manages DQA ESM Script and Configuration Pack to Local Machine

.DESCRIPTION

  Downloads DQA ESM Script and Configuration Pack to Local Machine

  1. Check if MD5 Hash has changed
  2. Bases configuation using DQAESM.config file
  3. Download Script and Configuration Pack from WebServer
  4. Unpack Script and Configuration Pack

.NOTES

  Version:        1.0.0.0
  Author:         Jean-Pierre Simonis - Delivery Quality Assurance (DQA)
  Creation Date:  20200309
  Purpose/Change: Initial Release

.EXAMPLE

  .\Sync-DQAESMSignature.ps1

#>

###########################
#    ESM Configuration    #
###########################

#Current Script Location
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

#ESM Configuration File
$ESMConfig = "$PSScriptRoot\DQAESM.config"
$ESMConfig = Get-Content -Raw -Path $ESMConfig | ConvertFrom-Json


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


#  Collect Configuration  #

### Set Base Variable
$UpdateESMPackage = $false

### Script Package Names
$ESMPackage = $ESMConfig.ESMPackage.Split(".")
$ESMPackageMD5 = $ESMPackage[0] + ".md5"
$ESMPackage = $ESMPackage[0] + ".zip"

### Script Package URLs
$ESMPackageMD5URL = $ESMConfig.ESMURL + "$ESMPackageMD5"
$ESMPackageURL = $ESMConfig.ESMURL + "$ESMPackage"

#      Version Check      #

### Check if MD5 exists for ESM Script Package

$localScriptPackMD5 = "$PSScriptRoot\$ESMPackageMD5"
$localScriptPack = "$PSScriptRoot\$ESMPackage"

$CheckMD5 = Test-Path -Path $localScriptPackMD5

### Clean installation of Script Pack (No previous download)
If ($CheckMD5 -eq $false) {

   Invoke-WebRequest -Method GET -Uri $ESMPackageMD5URL -OutFile $localScriptPackMD5 -UseBasicParsing
   Invoke-WebRequest -Method GET -Uri $ESMPackageURL -OutFile $localScriptPack -UseBasicParsing

   ###Set Flag to let script know to clear previous Outlook Email Signature from User Profile before unpacking new ESM Script pack
   $UpdateESMPackage = $true

} else {

   ###Collect and store local Script pack MD5 hash variable
   $localScriptPackMD5Hash = Get-Content -Raw -Path $localScriptPackMD5
   ###Collect and store remote Script pack MD5 hash variable
   $remoteScriptPackMD5Hash = Invoke-WebRequest -Method GET -Uri $ESMPackageMD5URL -UseBasicParsing

   ###Compare Hash Values if different overwrite local MD5 and Script Pack ZIP file otherwise do nothing
   $remoteScriptPackMD5Hash = [string]$remoteScriptPackMD5Hash
         If ($localScriptPackMD5Hash -ne $remoteScriptPackMD5Hash)
            {
               Write-Output "Downloading Latest ESM Script Package and MD5 Hash"
               #Download Latest Script Package MD5 and ZIP files
               Invoke-WebRequest -Method GET -Uri $ESMPackageMD5URL -OutFile $localScriptPackMD5 -UseBasicParsing
               Invoke-WebRequest -Method GET -Uri $ESMPackageURL -OutFile $localScriptPack -UseBasicParsing

               #Set Flag to let script know to clear previous Outlook Email Signature from User Profile before unpacking new ESM Script pack
               $UpdateESMPackage = $true

         } else {

            Write-Output "You have the latest ESM Script Pack"

         }
}


# Unzip ESM Script Package #

If ($UpdateESMPackage -eq $true){


   ### Extract Script Pack to Updater Script Execution Path
   $ExtractScriptPack = Unzip $localScriptPack $PSScriptRoot


}

