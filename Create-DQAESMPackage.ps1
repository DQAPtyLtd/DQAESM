<#
.SYNOPSIS

  Generates Web Server Deployment Package for DQA ESM Scripts and Configuration

.DESCRIPTION

  Generates Zip file and MD5 file for WEb Deployment of DQA ESM Script and Configuration Package

  1. Create Zip file Script Package and Template Package Management Scripts and DQA ESM Configuration file
  2. Create MD5 file of Script and Configuration Package

.NOTES

  Version:        1.0.0.0
  Author:         Jean-Pierre Simonis - Delivery Quality Assurance (DQA)
  Creation Date:  20200309
  Purpose/Change: Initial Release

.EXAMPLE

  .\Create-DQAESMPackage.ps1

#>

###########################
#    ESM Configuration    #
###########################

#Current Script Location
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition


#ESM Configuration File
$ESMConfig = "$PSScriptRoot\DQAESM.config"
$ESMConfig = Get-Content -Raw -Path $ESMConfig | ConvertFrom-Json

### Script Package Names
$ESMPackage = $ESMConfig.ESMPackage.Split(".")
$ESMPackageMD5 = $ESMPackage[0] + ".md5"
$ESMPackage = $ESMPackage[0] + ".zip"


###########################
#    Script Execution     #
###########################

#Create Script Package
Compress-Archive -LiteralPath "$PSScriptRoot\Update-DQAESM.ps1", "$PSScriptRoot\Sync-DQAESMSignature.ps1", "$PSScriptRoot\DQAESM.config", "$PSScriptRoot\run-dqaesmupdate.vbs", "$PSScriptRoot\run-dqaesmsigupdate.vbs" -DestinationPath "$PSScriptRoot\Packages\$ESMPackage" -Force

#Create MD5 Hash of Archive
$FileHash = Get-FileHash -Path "$PSScriptRoot\Packages\$ESMPackage"
#Write MD5 to Disk
Set-Content -Path "$PSScriptRoot\Packages\$ESMPackageMD5" -Value $FileHash.Hash -Force