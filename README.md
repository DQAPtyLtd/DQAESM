DQA Email Signature Manager
===========================

A solution designed to centrally deploy client-side Outlook Email Signatures and
update the signature block dynamically from Azure Active Directory attributes.

Solution Files and Purpose
--------------------------

-   **Generate-DQAESMConfig.ps1** – Generates DQAESM.config in JSON format used
    in client script package to provide client-side configuration

-   **Create-DQAESMPackage.ps1** – Generates DQAEMS.zip for client-side script
    deployment

-   **Create-DQAESMTemplatePackage.ps1** – Generates DQAESMSignatures.zip for
    client-side signature block templates

-   **Update-DQAESM.ps1** – Client-Side PowerShell Script that is packaged into
    DQAESM.zip that is configured as a scheduled task to manage script and
    configuration updates

-   **Sync-DQAESMSignature.ps1** – Client-Side PowerShell Script that is
    packaged into DQAESM.zip that is configured as a scheduled task to manage
    updates to the users email signature block

-   **run-dqaesmupdate.vbs** – VBScript wrapper that executes the
    Update-DQAESM.ps1 powershell script so that it doesnt popup interactively

-   **run-dqaesmsigupdate.vbs** – VBScript wrapper that executes the
    Sync-DQAESMSignature.ps1 powershell script so that it doesnt popup interactively

-   **Deploy-DQAESM.ps1** – PowerShell script designed to be deployed via Intune
    that installs the DQA ESM solution (designed to be in the end-user’s context
    not as system)

-   **DQAESM.config** – Client-Side master configuration file

-   **DQAESM-OutlookConfiguredSignatures.config** – Client-Side configuration
    file to track and compare the esm configured default email signatures in
    outlook (dynamically generated)

-   **DQAESM-UserProfile.config** – Client-Side configuration file to track and
    compare collected Azure AD user profile information against previously
    configured user profile information (dynamically generated)

-   **DQAESM.zip** – Script and Configuration Client Deployment Package (Name
    configurable via DQAESM.config)

    -   Contains: Update-DQAESM.ps1, Sync-DQAESMSignature.ps1 and DQAESM.config

-   **DQAESM.md5** – file that contains MD5 hash signature for DQAESM.zip it is
    used to compare with locally cached md5 file on the client computer if
    different the zip file is downloaded and scripts and configuration are
    updated on the client PC

-   **DQAESMSignatures.zip** – Signature Block Templates Client Deployment
    Package (Name configurable via DQAESM.config)

-   **DQAESMSignatures.md5** – file that contains MD5 hash signature for
    DQAESMSignatures.zip it is used to compare with locally cached md5 file on
    the client computer if different the zip file is downloaded and the email
    templates are updated on the client PC

Solution Folders and Purpose
----------------------------

### Solution Side

-   **Root Folder** – Holds all executable scripts and configuration files

-   **/Templates** – Location to place Email Signature Templates for client
    packaging

-   **/Packages** – Location that Client Deployment Packages are built into

### Client-Side

-   **Root Folder/Installation Folder –** Defaults to **%appdata%\\DQAESM also**
    holds all executable scripts and configuration files

-   **/Templates** - Location to place downloaded Email Signature Client
    Deployment Packages

Installation
------------

### Requirements

-   Web Server accessible by end users for hosting of Client Deployment Packages

-   Azure Active Directory

-   Azure Active Directory App Registration with Microsoft Graph Permissions to
    allow solution scripts to read the necessary user profile information

-   Microsoft Intune

-   Windows 10

-   Microsoft Outlook 2013 and above

### Hosting

This solution requires a webserver accessible to your end users to host four (4)
files

-   **DQAESM.zip** (Configurable) – Scripts and Client-Side Configuration

-   **DQAESM.md5** – File Hash Signature of DQAESM.zip used for update
    management

-   **DQAESMSignatures.zip** (Configurable) – Outlook Signature Block Templates

-   **DQAESMSignatures.md5** – File Hash Signature of DQAESMSignatures.zip used
    for update management

### Azure Active Directory App Registration

In a future iteration of this solution this will be automated.

As an Azure AD administrator from the Azure Portal (https://portal.azure.com):

-   Create an App Registration for this Solution eg. **DQA Email Signature
    Manager**

-   Take note of the **AppID**

-   Also ensure that you grant application api permissions to **read** from
    **Microsoft Graph** the necessary user profile attributes eg.
    **User.Read.All**

-   Create and take note of the **client secret** for this App Registration

-   Also take not of the **Directory ID** under your Azure AD instance (under
    Properties blade)

### Generate solution configuration

1.  Open **Generate-DQAESMConfig.ps1** in a PowerShell Editor like Visual Studio
    Code

2.  Modify the solution configuration

| Description                                                                                                      | Configuration Item  | Valid Values                                                                                                                                                                                                                                                |
|------------------------------------------------------------------------------------------------------------------|---------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| Force Default Outlook Default Signature                                                                          | SetOutlookDefaults  | True or False                                                                                                                                                                                                                                               |
| Name of Email Signature Template required for New emails                                                         | OutlookDefaultNew   | Free Text (Name must match what is visible in outlook)                                                                                                                                                                                                      |
| Name of Email Signature Template required for reply/forwarded emails                                             | OutlookDefaultReply | Free Text (Name must match what is visible in outlook)                                                                                                                                                                                                      |
| DQA ESM Email Signature Template Package name                                                                    | ESMTemplatePackage  | Free Text (Desired name of zip file eg. DQAESMSignatures.zip)                                                                                                                                                                                               |
| DQA ESM Email Signature Prefix                                                                                   | ESMSignaturePrefix  | Free Text (Prefix is used to identity signature deployed by this tool when you create your email templates ensure your configured prefix matches the what is configured in outlook signature name eg. DQAESM-\<your templatename\> refer to sample content) |
| DQA ESM Script and Configuration Package name                                                                    | ESMPackage          | Free Text (Desired name of zip file eg. DQAESM.zip)                                                                                                                                                                                                         |
| DQA ESM Script and Configuration Package base url                                                                | ESMURL              | Free Text (base url of script package file eg. <https://sampleurl.com/esm/>)                                                                                                                                                                                |
| DQA ESM Email Signature Template Package base url **Note:** url does not have to different to ESM Script Package | ESMTemplatesURL     | Free Text (base url of template package file eg. <https://sampleurl.com/signatures/>)                                                                                                                                                                       |
| DQA ESM Signature Block Defaults. These defaults are used with Azure AD value is null or empty                   | ESMTemplateDefaults | Free Text (Example content are the only fields currently support by this solution without modify the code in sync-dqaesmsignature.ps1)                                                                                                                      |
| Azure Active Directory Solution App Registration                                                                 | ESMClientID         | Free Text (AppID noted from Azure AD App Registration step)                                                                                                                                                                                                 |
| Azure Active Directory Directory ID                                                                              | ESMTenantID         | Free Text (Directory ID noted from Azure AD App Registration step)                                                                                                                                                                                          |
| Azure Active Directory Solution App Registration Client Secret (authentication key)                              | ESMclientSecret     | Free Text (Client Secret noted from Azure AD App Registration step)                                                                                                                                                                                         |

3.  Once PowerShell configuration saved execute this script

4.  You will now have **DQAESM.config** file in the same directory as the
    generate-dqaesmconfig.ps1

### Create Outlook Signature templates

1.  Create your email signature just as you would if you were creating one for
    yourself in Outlook

    1.  Note the following:

        1.  Ensure you have the same prefix for signatures you want to make into
            a template eg. DQAESM-

        2.  You have the ability to have separate or the same template for New
            emails vs Reply/Forwarded email

        3.  Use the following text placeholders in your signature as the
            solution will replace these values with Azure AD values (only the
            follow values are supported without the need to modify the code in
            sync-dqaesmsignature.ps1 refer to example content

            1.  **DQAFULLNAME** – Azure AD Display Name

            2.  **DQATITLE** – Azure AD Job Title

            3.  **DQAMOBILE** – Azure AD Mobile Phone

            4.  **DQAADDRESS** – Azure AD Street Address

            5.  **DQACITY** – Azure AD City

            6.  **DQASTATE** – Azure AD State

            7.  **DQAPOSTCODE** – Azure AD Postal Code

2.  Once you have created your email signatures in Outlook. Open File Explorer
    to the following path **%appdata%\\Microsoft\\Signatures**

3.  Copy the files and folders with names that match the Email Signature
    templates you have just created to the Templates folder of this cloned
    solution

### Create Deployment Packages

1.  Execute the **Create-DQAESMPackage.ps1** script

2.  You will now have a **DQAESM.zip** and **DQAESM.md5** in the packages folder
    (**Note:** The file names will be based on what is configured in the
    DQAESM.config file)

3.  Execute the **Create-DQAESMTemplatePackage.ps1** script

4.  You will now have a **DQAESMSignatures.zip** and **DQAESMSignatures.md5** in
    the packages folder (**Note:** The file names will be based on what is
    configured in the DQAESM.config file)

5.  Upload the zip and md5 files to the locations specified in the
    **DQAESM.config** file

### Client Deployment

1.  Open **Deploy-DQAESM.ps1** in a PowerShell Editor like Visual Studio Code

2.  Modify the solution configuration under ESM Configuration section of this
    script to match the settings configured in the **DQAESM.config** file

3.  From the Azure Portal use Intune to deploy the **Deploy-DQAESM.ps1** script
    file

    1.  Under **Intune** from the **Device Configuration** Blade

    2.  Open the **Scripts** Blade

    3.  Click **+ Add**

    4.  Supply a Name eg. Deploy DQQA Email Signature Manager

    5.  Upload/Select the **Deploy-DQAESM.ps1** from your computer

    6.  Ensure **Run this script using the logged on credentials** is set to
        **yes**

    7.  Assign a deployment group

Updating the Configuration/Script Package
-----------------------------------------

1.  Follow the instructions for **generating solution configuration** and/or
    make any required changes to solution scripts

2.  Follow Steps **1** and **2** and **5** from the **Create Deployment
    Packages** instructions

Updating the Email Templates Package
------------------------------------

1.  Follow the instructions for **Create Outlook Signature Templates**

2.  Follow Steps **3** and **4** and **5** from the **Create Deployment
    Packages** instructions
