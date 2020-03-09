<#
.SYNOPSIS

  Generates Web Server Deployment Package for DQA ESM Signature Templates

.DESCRIPTION

  Generates Zip file and MD5 file for WEb Deployment of DQA ESM Signature Templates

  1. Create Zip file Signature Block Templates
  2. Create MD5 file of Signature Block Package

.NOTES

  Version:        1.0.0.0
  Author:         Jean-Pierre Simonis - Delivery Quality Assurance (DQA)
  Creation Date:  20200309
  Purpose/Change: Initial Release

.EXAMPLE

  .\Create-DQAESMTemplatePackage.ps1

#>

###########################
#    ESM Configuration    #
###########################

#Current Script Location
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

#ESM Configuration File
$ESMConfig = "$PSScriptRoot\DQAESM.config"
$ESMConfig = Get-Content -Raw -Path $ESMConfig | ConvertFrom-Json

### Template Package Names
$ESMTemplatePackage = $ESMConfig.ESMTemplatePackage.Split(".")
$ESMTemplatePackageMD5 = $ESMTemplatePackage[0] + ".md5"
$ESMTemplatePackage = $ESMTemplatePackage[0] + ".zip"

###########################
#    Script Execution     #
###########################

#Create Script Package
$Templates = "$PSScriptRoot\Templates"
& chdir $Templates
Compress-Archive -Path * -DestinationPath "$PSScriptRoot\Packages\$ESMTemplatePackage" -Force

#Create MD5 Hash of Archive
$FileHash = Get-FileHash -Path "$PSScriptRoot\Packages\$ESMTemplatePackage"
#Write MD5 to Disk
Set-Content -Path "$PSScriptRoot\Packages\$ESMTemplatePackageMD5" -Value $FileHash.Hash -Force