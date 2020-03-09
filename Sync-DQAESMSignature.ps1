<#
.SYNOPSIS

  Synchronises and manages DQA ESM Signatures Templates Packs to Local Machine and applies current user information to signature block based on Azure AD information

.DESCRIPTION

  Downloads DQA ESM Template Packs and Updates Outlook Signature Blocks based on Azure AD information of current user

  1. Check if MD5 Hash has changed
  2. Bases configuation using DQAESM.config file
  3. Download Signature Pack from WebServer
  4. Unpack Template Pack
  5. Download Azure AD User information
  6. Update Signature Block with user information
  7. Update Outlook Signatures


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

#ESM User Profile
$ESMUserProfilePath = "$PSScriptRoot\DQAESM-UserProfile.config"
$ESMUserProfile = Get-Content -Raw -Path $ESMUserProfilePath -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue

#Outlook Configured Signatures
$OutlookCfgSignaturesPath = "$PSScriptRoot\DQAESM-OutlookConfiguredSignatures.config"
$OutlookCfgSignatures = Get-Content -Raw -Path $OutlookCfgSignaturesPath -ErrorAction SilentlyContinue | ConvertFrom-Json -ErrorAction SilentlyContinue

#Outlook Signature Path
$SignaturePath = "$env:userprofile\AppData\Roaming\Microsoft\Signatures"


###########################
#       Functions         #
###########################

#function to set default outlook email signature
function Set-OutlookDefaultEmailSignatures {
   Param(
      $newsignature,
      $replysignature,
      $outFilePath
  )

   #Open Outlook COM Object to determine Office Version
   $objoutlook = new-object -comobject outlook.application
   #Split the retrieved version number with full stops as a delimiter
   $outlookVersion = $objoutlook.version.Split(".")
   #Get the first array member (MS OFfice Major Version) and append a .0
   $outlookVersion = $outlookVersion[0] + ".0"
   #Close Outlook Com object
   $objoutlook.quit()

   #Set Outlook Signature Settings
   $SetNewEmailSignature = New-ItemProperty -Path HKCU:\Software\Microsoft\Office\$outlookVersion\Common\MailSettings -name "NewSignature" -PropertyType ExpandString -Value $newsignature -ErrorAction SilentlyContinue
   $SetReplyEmailSignature = New-ItemProperty -Path HKCU:\Software\Microsoft\Office\$outlookVersion\Common\MailSettings -name "ReplySignature" -PropertyType ExpandString -Value $replysignature -ErrorAction SilentlyContinue

   #JSON Body response
   $outlookconfiguredSignaturesParameters = @{

      OutlookDefaultNew = "$newsignature"
      OutlookDefaultReply = "$replysignature"

   }

   #Prepare Response
   $outlookconfiguredSignaturesParameters = ConvertTo-Json $outlookconfiguredSignaturesParameters

   #Write to Disk
   Set-Content -Path $outFilePath -Value $outlookconfiguredSignaturesParameters

}

#function to remove default outlook email signature setting
function Remove-OutlookDefaultEmailSignatures {


   #Open Outlook COM Object to determine Office Version
   $objoutlook = new-object -comobject outlook.application
   #Split the retrieved version number with full stops as a delimiter
   [string]$outlookVersion = $objoutlook.version.Split(".")
   #Get the first array member (MS OFfice Major Version) and append a .0
   [string]$outlookVersion = $outlookVersion[0] + ".0"
   #Close Outlook Com object
   $objoutlook.quit()

   #Set Outlook Signature Settings
   $NewEmailSignature = Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\$outlookVersion\Common\MailSettings -name "NewSignature" -ErrorAction SilentlyContinue
   $ReplyEmailSignature = Remove-ItemProperty -Path HKCU:\Software\Microsoft\Office\$outlookVersion\Common\MailSettings -name "ReplySignature" -ErrorAction SilentlyContinue

}

#Function to Collect User information from Azure AD Graph
function Get-UserProfileInfo {
   Param(
      $clientID,
      $tenantId,
      $clientSecret,
      $UserEmail
  )
   # Azure AD OAuth Application Token for Graph API
   # Get OAuth token for a AAD Application (returned as $token)
   # Construct URI
   $graphTokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

   # Construct Body
   $body = @{
      client_id     = $clientId
      scope         = "https://graph.microsoft.com/.default"
      client_secret = $clientSecret
      grant_type    = "client_credentials"
   }

   # Get OAuth 2.0 Token
   $tokenRequest = Invoke-WebRequest -Method Post -Uri $graphTokenUri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing

   # Access Token
   $token = ($tokenRequest.Content | ConvertFrom-Json).access_token

   # Graph API call in PowerShell using obtained OAuth token (see other gists for more details)

   # Specify the URI to call and method
   $userInfoUri = "https://graph.microsoft.com/v1.0/users/" + $UserEmail + "?`$select=id,userPrincipalName,displayName,jobTitle,givenName,surname,department,mail,city,officeLocation,mobilePhone,businessPhones,streetaddress,postalcode,state,country,companyname"

   #Get Users Manager
   #$userInfoUri = "https://graph.microsoft.com/v1.0/users/" + $UserEmail + "/manager"
   #Get Users Direct Reports
   #$userInfoUri = "https://graph.microsoft.com/v1.0/users/" + $UserEmail + "/directreports"
   #Get Users Photo
   #$userInfoUri = "https://graph.microsoft.com/v1.0/users/" + $UserEmail + "/photos/96x96/`$value"

   $method = "GET"

   # Run Graph API query
   $query = Invoke-WebRequest -Method $method -Uri $userInfoUri -ContentType "application/json" -Headers @{Authorization = "Bearer $token"} -ErrorAction Stop
   $FunctionReturn = ($query.Content | ConvertFrom-Json)

   #Return Collected Information
   Return $FunctionReturn

}

#Function to extract zip files
function unzip ($file,$ExtractPath) {
   expand-archive -Path $file -DestinationPath $ExtractPath -Force
}


###########################
#    Script Execution     #
###########################


#  Collect Configuration  #

### Application (client) ID, tenant ID and secret
$clientId = $ESMConfig.ESMclientId
$tenantId = $ESMConfig.ESMtenantId
$clientSecret = $ESMConfig.ESMclientSecret

### Template Package Names
$ESMTemplatePackage = $ESMConfig.ESMTemplatePackage.Split(".")
$ESMTemplatePackageMD5 = $ESMTemplatePackage[0] + ".md5"
$ESMTemplatePackage = $ESMTemplatePackage[0] + ".zip"

#TemplatePackPrefix
$SignaturePrefix = $ESMConfig.ESMSignaturePrefix

### Template Package URLs
$ESMTemplatePackageMD5URL = $ESMConfig.ESMTemplatesURL + "$ESMTemplatePackageMD5"
$ESMTemplatePackageURL = $ESMConfig.ESMTemplatesURL + "$ESMTemplatePackage"

### Signature Defaults if Azure AD values are null/empty

$DefaultJobTitle = $ESMConfig.ESMTemplateDefaults.JobTitle
$DefaultPhone = $ESMConfig.ESMTemplateDefaults.BusinessPhone
$DefaultAddress = $ESMConfig.ESMTemplateDefaults.Address
$DefaultCity = $ESMConfig.ESMTemplateDefaults.City
$DefaultState = $ESMConfig.ESMTemplateDefaults.State
$DefaultPostCode = $ESMConfig.ESMTemplateDefaults.PostCode

### Configure Outlook Default Signatures

$SetOutlookDefaults = $ESMConfig.SetOutlookDefaults
$OutlookNewSignature = $ESMConfig.OutlookDefaultNew
$OutlookReplySignature = $ESMConfig.OutlookDefaultReply

#  Collect User Details   #

### Get Current user from Windows Session
$CurrentUser = &whoami /upn
### Collect current users Azure AD information
$Profile = Get-UserProfileInfo -clientID $clientId -tenantId $tenantId -clientSecret $clientSecret -UserEmail $CurrentUser

Write-Output "Collected User Profile Information"
Write-Output $Profile

### Update user values with Default values if they are null or empty

IF ($Profile.jobTitle -eq $null -or $Profile.jobTitle -eq "") { $Profile.jobTitle = $DefaultJobTitle }
IF ($Profile.mobilePhone -eq $null -or $Profile.mobilePhone -eq "") { $Profile.mobilePhone = $DefaultPhone }
IF ($Profile.streetAddress -eq $null -or $Profile.streetAddress -eq "") { $Profile.streetAddress = $DefaultAddress }
IF ($Profile.city -eq $null -or $Profile.city -eq "") { $Profile.city = $DefaultCity }
IF ($Profile.state -eq $null -or $Profile.state -eq "") { $Profile.state = $DefaultState }
IF ($Profile.postalCode -eq $null -or $Profile.postalCode -eq "") { $Profile.postalCode = $DefaultPostCode }

#      Version Check      #


### Ensure Template Directory Exists
$templateDestinationFolder = "$PSScriptRoot\Templates"

# Create Program Files Folder
if (-not(test-path $templateDestinationFolder)) {
   try {

       $null = New-Item -type "Directory" -Path $templateDestinationFolder -Force

   }
   catch {
       Write-Output "Failed to create program folder $templateDestinationFolder"
       Exit 911
   }
}

### Check if MD5 exists for ESM Template Package

$localTemplatePackMD5 = "$templateDestinationFolder\$ESMTemplatePackageMD5"
$localTemplatePack = "$templateDestinationFolder\$ESMTemplatePackage"

$CheckMD5 = Test-Path -Path $localTemplatePackMD5

### Clean installation of Template Pack (No previous download)
If ($CheckMD5 -eq $false) {

   Invoke-WebRequest -Method GET -Uri $ESMTemplatePackageMD5URL -OutFile $localTemplatePackMD5
   Invoke-WebRequest -Method GET -Uri $ESMTemplatePackageURL -OutFile $localTemplatePack

   ###Set Flag to let script know to clear previous Outlook Email Signature from User Profile before unpacking new ESM Template pack
   $UpdateESMTemplate = $true

} else {

   ###Collect and store local Template pack MD5 hash variable
   $localTemplatePackMD5Hash = Get-Content -Raw -Path $localTemplatePackMD5
   ###Collect and store remote Template pack MD5 hash variable
   $remoteTemplatePackMD5Hash = Invoke-WebRequest -Method GET -Uri $ESMTemplatePackageMD5URL

   ###Compare Hash Values if different overwrite local MD5 and Template Pack ZIP file otherwise do nothing
   $remoteTemplatePackMD5Hash = [string]$remoteTemplatePackMD5Hash
         If ($localTemplatePackMD5Hash -ne $remoteTemplatePackMD5Hash)
            {
               Write-Output "Downloading Latest ESM Template Package and MD5 Hash"
               #Download Latest Template Package MD5 and ZIP files
               Invoke-WebRequest -Method GET -Uri $ESMTemplatePackageMD5URL -OutFile $localTemplatePackMD5
               Invoke-WebRequest -Method GET -Uri $ESMTemplatePackageURL -OutFile $localTemplatePack

               #Set Flag to let script know to clear previous Outlook Email Signature from User Profile before unpacking new ESM Template pack
               $UpdateESMTemplate = $true

         } else {

            Write-Output "You have the latest ESM Template Pack"

         }
}

# Store User Profile Details #

Write-Output "User Profile Information (Updated with Default Values if empty)"
Write-Output $Profile

### Convert updated User Profile details to JSON for later comparison (in case Azure AD attributes get updated)
$convertUserProfiletoJSON = ConvertTo-Json $Profile

### Write user profile details to Disk
Set-Content -Path $ESMUserProfilePath -Value $convertUserProfiletoJSON

# Compare Azure AD User Profile to Cached user profile details #

### Compare Array objects current vs local cache

$compareLocalProfile = $esMuserprofile | convertto-json -Compress -Depth 50
$compareCloudProfile = $profile | convertto-json -Compress -Depth 50

### If Array objects do not match then the Cloud based user attributes have changed and the email signature will be redeployed
if ($compareLocalProfile -ne $compareCloudProfile) {

   $UpdateESMTemplate = $true

}

# Clear Existing ESM Templates #

If ($UpdateESMTemplate -eq $true){

   ### Define ESM Templates to Delete (ONLY DELETE ESM Templates)
   $ClearExistingSignaurePath = $SignaturePath + "\" + $SignaturePrefix + "*"

   ### Delete Outlook Signatures
   Remove-Item -path $ClearExistingSignaurePath -Recurse -force

   ### Extract Template Pack to Outlook Signature Path
   $ExtractTemplatePack = Unzip $localTemplatePack $SignaturePath

   ### Update Template Content with AZure AD Info

   ## ESM Template file filters
   $htmFiles = $SignaturePrefix + "*.htm"
   $rtfFiles = $SignaturePrefix + "*.rtf"
   $txtFiles = $SignaturePrefix + "*.txt"

   ## Collection of ESM Template Pack files
   $colHtmFile = Get-ChildItem $SignaturePath -Filter $htmFiles
   $colRtfFile = Get-ChildItem $SignaturePath -Filter $rtfFiles
   $colTxtFile = Get-ChildItem $SignaturePath -Filter $txtFiles

   ## Loop through all htm files and update content with Azure AD information

   Foreach ($file in $colHtmFile)
   {

      $filepath = $file.FullName
      $content = Get-Content $filepath -raw
      $content = $content -replace 'DQAFULLNAME', $Profile.displayName
      $content = $content -replace 'DQATITLE', $Profile.jobTitle
      $content = $content -replace 'DQAMOBILE', $Profile.mobilePhone
      $content = $content -replace 'DQAADDRESS', $Profile.streetAddress
      $content = $content -replace 'DQACITY', $Profile.city
      $content = $content -replace 'DQASTATE', $Profile.state
      $content = $content -replace 'DQAPOSTCODE', $Profile.postalCode
      $content | Set-Content $filepath

   }

   ## Loop through all rtf files and update content with Azure AD information

   Foreach ($file in $colRtfFile)
   {

      $filepath = $file.FullName
      $content = Get-Content $filepath -raw
      $content = $content -replace 'DQAFULLNAME', $Profile.displayName
      $content = $content -replace 'DQATITLE', $Profile.jobTitle
      $content = $content -replace 'DQAMOBILE', $Profile.mobilePhone
      $content = $content -replace 'DQAADDRESS', $Profile.streetAddress
      $content = $content -replace 'DQACITY', $Profile.city
      $content = $content -replace 'DQASTATE', $Profile.state
      $content = $content -replace 'DQAPOSTCODE', $Profile.postalCode
      $content | Set-Content $filepath

   }

   ## Loop through all txt files and update content with Azure AD information

   Foreach ($file in $colTxtFile)
   {

      $filepath = $file.FullName
      $content = Get-Content $filepath -raw
      $content = $content -replace 'DQAFULLNAME', $Profile.displayName
      $content = $content -replace 'DQATITLE', $Profile.jobTitle
      $content = $content -replace 'DQAMOBILE', $Profile.mobilePhone
      $content = $content -replace 'DQAADDRESS', $Profile.streetAddress
      $content = $content -replace 'DQACITY', $Profile.city
      $content = $content -replace 'DQASTATE', $Profile.state
      $content = $content -replace 'DQAPOSTCODE', $Profile.postalCode
      $content | Set-Content $filepath

   }

}

#  Set Default Email Signature for Outlook  #
if ($SetOutlookDefaults -eq $true) {

   if ($OutlookNewSignature -eq $OutlookCfgSignatures.OutlookDefaultNew -and $OutlookReplySignature -eq $OutlookCfgSignatures.OutlookDefaultReply ) {#Do Nothing
   } else {
      $ConfigureOutlook = Set-OutlookDefaultEmailSignatures -newsignature $OutlookNewSignature -replysignature $OutlookReplySignature -outFilePath $OutlookCfgSignaturesPath
   }

} else {

   $ConfigureOutlook = Remove-OutlookDefaultEmailSignatures

}