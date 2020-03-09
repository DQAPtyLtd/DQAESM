<#
.SYNOPSIS

  Generates Configuration for DQA ESM

.DESCRIPTION

  Installs the DQA ESM Program

  1. Generates DQAESM.config in JSON format the same directory where this script is executed
  2. Bases configuation on values configured in this script


.NOTES

  Version:        1.0.0.0
  Author:         Jean-Pierre Simonis - Delivery Quality Assurance (DQA)
  Creation Date:  20200309
  Purpose/Change: Initial Release

.EXAMPLE

# Update configuration values within the script then execute below
  .\Generate-DQAESMConfig.ps1


#>

###########################
#    ESM Configuration    #
###########################

 #JSON Output path
    #Current Script Location
    $PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

    $outFilePath = "$PSScriptRoot\DQAESM.config"

 #JSON Body response (Update values below to customise your DQA ESM Deployment)
    $configureEmailSignatureManagerParameters= @{
        SetOutlookDefaults = "True"
        OutlookDefaultNew = "DQAESM-New-MS"
        OutlookDefaultReply = "DQAESM-Reply-Forward"
        ESMTemplatePackage = "DQAESMSignatures.zip"
        ESMSignaturePrefix = "DQAESM-"
        ESMPackage = "DQAESM.zip"
        ESMURL = "https://sampleurl.com/esm/"
        ESMTemplatesURL = "https://sampleurl.com/signatures/"

        ESMTemplateDefaults = @{

            "JobTitle" = "Delivery Specialist"
            "BusinessPhone" = "1800 111 111"
            "Address"= "123 Seasame Street"
            "State" = "VIC"
            "City" = "Melbourne"
            "PostCode" = "3000"
        }
        ESMclientId = "Insert AzureAD App ID"
        ESMtenantId = "Insert AzureAD Tenant ID"
        ESMclientSecret = "Insert AzureAD Client Secret for registered app"

    }

###########################
#    Script Execution     #
###########################

#Prepare Response
$configureEmailSignatureManagerParameters = ConvertTo-Json $configureEmailSignatureManagerParameters

#Write to Disk
Set-Content -Path $outFilePath -Value $configureEmailSignatureManagerParameters

#Show Output to Screen
Write-Output $configureEmailSignatureManagerParameters