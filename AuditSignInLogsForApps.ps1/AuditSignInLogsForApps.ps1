<#
.SYNOPSIS
    Retrieves Microsoft Graph sign-in logs for specified applications within a defined period and exports them to Excel.

.DESCRIPTION
    This script connects to Microsoft Graph using the 'AuditLog.Read.All' scope.
    It defines a list of application IDs and their corresponding friendly names.
    It then queries the Microsoft Graph sign-in logs for interactive, service principal,
    and managed identity sign-ins related to these applications that occurred within the specified number of past days.
    The results are combined, enriched with the application's friendly name, and exported to an Excel file.
    The Excel file includes a data sheet and a pivot table summarizing sign-ins by application and user.
    The export path and filename are configurable.

.PARAMETER DaysToQuery
    The number of past days for which to retrieve sign-in logs. Defaults to 3 days.

.PARAMETER ExportPath
    The directory path where the Excel report will be saved. Defaults to the user's temporary directory.

.NOTES
    Author: itsocjn
    Date: 29/04/2025
    Version: 1.1 - Added parameters, dynamic filename, configurable path, comments, and formatting.

.EXAMPLE
    .\AuditSignInLogsForMSApps.ps1
    Runs the script with default settings (3 days of logs, export to temp directory).

.EXAMPLE
    .\AuditSignInLogsForMSApps.ps1 -DaysToQuery 7 -ExportPath "C:\Reports\Audit"
    Runs the script to retrieve logs for the past 7 days and saves the Excel file to "C:\Reports\Audit".

.LINK
    Install-Module ImportExcel -Scope CurrentUser # Required for Export-Excel cmdlet
    Connect-MgGraph # Ensure you are connected before running or the script will prompt.
#>
param (
    # Specify the number of past days to query for sign-in logs
    [int]$DaysToQuery = 30,

    # Specify the directory path for the Excel export
    [string]$ExportPath = $env:TEMP
)

# Ensure the Export Path exists, create if it doesn't
if (-not (Test-Path -Path $ExportPath -PathType Container)) {
    Write-Warning "Export path '$ExportPath' not found. Attempting to create it."
    try {
        New-Item -Path $ExportPath -ItemType Directory -Force -ErrorAction Stop | Out-Null
        Write-Host "Successfully created export directory: $ExportPath"
    }
    catch {
        Write-Error "Failed to create export directory '$ExportPath'. Please check permissions or specify a different path. Error: $_"
        exit 1 # Exit the script if the directory cannot be created
    }
}

# --- Configuration ---

# Authenticate to Microsoft Graph
# Scopes required: AuditLog.Read.All
# If not already connected, this will prompt for authentication.
Connect-MgGraph -Scopes "AuditLog.Read.All" -ErrorAction SilentlyContinue # Use SilentlyContinue if you expect to be pre-connected often
if (-not (Get-MgContext)) {
    Write-Error "Failed to connect to Microsoft Graph. Please ensure the module is installed and you have permissions."
    exit 1
}
Write-Host "Successfully connected to Microsoft Graph."

# Define the application IDs and their corresponding names
# Add or remove applications as needed. Comment out lines with '#' to exclude them temporarily.
$applications = @{
    "12128f48-ec9e-42f0-b203-ea49fb6af367" = "Teams PowerShell Module"
    "1950a258-227b-4e31-a9cf-717495945fc2" = "Power Platform PowerShell Module" # Also used by Microsoft Azure PowerShell
    "9bc3ab49-b65d-410a-85ad-de819febfddc" = "SharePoint PowerShell Module"
    "fb78d390-0c51-40cd-8e17-fdbfab77341b" = "Exchange PowerShell Module" # Also used by Microsoft Exchange REST API Based Powershell
    "1b730954-1685-4b74-9bfd-dac224a7b894" = "Azure AD PowerShell Module" 
    # --- Other Microsoft Applications (Examples - uncomment or add as needed) ---
    #"392cab40-8474-4fa9-a108-9ce447bf8c18" = "Taegis"   
    #"23523755-3a2b-41ca-9315-f81f3f566a95" = "ACOM Azure Website"
    #"74658136-14ec-4630-ad9b-26e160ff0fc6" = "ADIbizaUX"
    #"69893ee3-dd10-4b1c-832d-4870354be3d8" = "AEM-DualAuth"
    #"7ab7862c-4c57-491e-8a45-d52a7e023983" = "App Service"
    #"0cb7b9ec-5336-483b-bc31-b15b5788de71" = "ASM Campaign Servicing"
    #"7b7531ad-5926-4f2d-8a1d-38495ad33e17" = "Azure Advanced Threat Protection"
    #"e9f49c6b-5ce5-44c8-925d-015017e9f7ad" = "Azure Data Lake"
    #"835b2a73-6e10-4aa5-a979-21dfda45231c" = "Azure Lab Services Portal"
    #"c44b4083-3bb0-49c1-b47d-974e53cbdf3c" = "Azure Portal"
    #"022907d3-0f1b-48f7-badc-1ba6abab6d66" = "Azure SQL Database"
    #"37182072-3c9c-4f6a-a4b3-b3f91cacffce" = "AzureSupportCenter"
    #"9ea1ad79-fdb6-4f9a-8bc3-2b70f96e34c7" = "Bing"
    #"20a11fe0-faa8-4df5-baf2-f965f8f9972e" = "ContactsInferencingEmailProcessor"
    #"bb2a2e3a-c5e7-4f0a-88e0-8e01fd3fc1f4" = "CPIM Service"
    #"e64aa8bc-8eb4-40e2-898b-cf261a25954f" = "CRM Power BI Integration"
    #"00000007-0000-0000-c000-000000000000" = "Dataverse"
    #"60c8bde5-3167-4f92-8fdb-059f6176dc0f" = "Enterprise Roaming and Backup"
    #"497effe9-df71-4043-a8bb-14cf78c4b63b" = "Exchange Admin Center"
    #"f5eaa862-7f08-448c-9c4e-f4047d4d4521" = "FindTime"
    #"b669c6ea-1adf-453f-b8bc-6d526592b419" = "Focused Inbox"
    #"c35cb2ba-f88b-4d15-aa9d-37bd443522e1" = "GroupsRemoteApiRestClient"
    #"d9b8ec3a-1e4e-4e08-b3c2-5baf00c0fcb0" = "HxService"
    #"a57aca87-cbc0-4f3c-8b9e-dc095fdc8978" = "IAM Supportability"
    #"16aeb910-ce68-41d1-9ac3-9e1673ac9575" = "IrisSelectionFrontDoor"
    #"d73f4b35-55c9-48c7-8b10-651f6f2acb2e" = "MCAPI Authorization Prod"
    #"944f0bd1-117b-4b1c-af26-804ed95e767e" = "Media Analysis and Transformation Service"
    #"0cd196ee-71bf-4fd6-a57c-b491ffd4fb1e" = "Microsoft 365 Security and Compliance Center"
    #"0000000c-0000-0000-c000-000000000000" = "Microsoft App Access Panel"
    #"65d91a3d-ab74-42e6-8a2f-0add61688c74" = "Microsoft Approval Management"
    #"38049638-cc2c-4cde-abe4-4479d721ed44" = "Microsoft Authentication Broker"
    #"04b07795-8ddb-461a-bbee-02f9e1bf7b46" = "Microsoft Azure CLI"
    #"1950a258-227b-4e31-a9cf-717495945fc2" = "Microsoft Azure PowerShell"
    #"0000001a-0000-0000-c000-000000000000" = "MicrosoftAzureActiveAuthn"
    #"cf36b471-5b44-428c-9ce7-313bf84528de" = "Microsoft Bing Search"
    #"2d7f3606-b07d-41d1-b9d2-0d0c9296a6e8" = "Microsoft Bing Search for Microsoft Edge"
    #"1786c5ed-9644-47b2-8aa0-7201292175b6" = "Microsoft Bing Default Search Engine"
    #"3090ab82-f1c1-4cdf-af2c-5d7a6f3e2cc7" = "Microsoft Defender for Cloud Apps"
    #"60ca1954-583c-4d1f-86de-39d835f3e452" = "Microsoft Defender for Identity (formerly Radius Aad Syncer)"
    #"18fbca16-2224-45f6-85b0-f7bf2b39b3f3" = "Microsoft Docs"
    #"00000015-0000-0000-c000-000000000000" = "Microsoft Dynamics ERP"
    #"6253bca8-faf2-4587-8f2f-b056d80998a7" = "Microsoft Edge Insider Addons Prod"
    #"99b904fd-a1fe-455c-b86c-2f9fb1da7687" = "Microsoft Exchange ForwardSync"
    #"00000007-0000-0ff1-ce00-000000000000" = "Microsoft Exchange Online Protection"
    #"fb78d390-0c51-40cd-8e17-fdbfab77341b" = "Microsoft Exchange REST API Based Powershell"
    #"47629505-c2b6-4a80-adb1-9b3a3d233b7b" = "Microsoft Exchange Web Services"
    #"6326e366-9d6d-4c70-b22a-34c7ea72d73d" = "Microsoft Exchange Message Tracking Service"
    #"c9a559d2-7aab-4f13-a6ed-e7e9c52aec87" = "Microsoft Forms"
    #"00000003-0000-0000-c000-000000000000" = "Microsoft Graph"
    #"74bcdadc-2fdc-4bb3-8459-76d06952a0e9" = "Microsoft Intune Web Company Portal"
    #"fc0f3af4-6835-4174-b806-f7db311fd2f3" = "Microsoft Intune Windows Agent"
    #"d3590ed6-52b3-4102-aeff-aad2292ab01c" = "Microsoft Office"
    #"00000006-0000-0ff1-ce00-000000000000" = "Microsoft Office 365 Portal"
    #"67e3df25-268a-4324-a550-0de1c7f97287" = "Microsoft Office Web Apps Service"
    #"d176f6e7-38e5-40c9-8a78-3998aab820e7" = "Microsoft Online Syndication Partner Portal"
    #"5d661950-3475-41cd-a2c3-d671a3162bc1" = "Microsoft Outlook"
    #"93625bc8-bfe2-437a-97e0-3d0060024faa" = "Microsoft password reset service"
    #"871c010f-5e61-4fb1-83ac-98610a7e9110" = "Microsoft Power BI"
    #"28b567f6-162c-4f54-99a0-6887f387bbcc" = "Microsoft Storefronts"
    #"cf53fce8-def6-4aeb-8d30-b158e7b1cf83" = "Microsoft Stream Portal"
    #"98db8bd6-0cc0-4e67-9de5-f187f1cd1b41" = "Microsoft Substrate Management"
    #"fdf9885b-dd37-42bf-82e5-c3129ef5a302" = "Microsoft Support"
    #"1fec8e78-bce4-4aaf-ab1b-5451cc387264" = "Microsoft Teams"
    #"cc15fd57-2c6c-4117-a88c-83b1d56b4bbe" = "Microsoft Teams Services"
    #"5e3ce6c0-2b1f-4285-8d4b-75ee78787346" = "Microsoft Teams Web Client"
    #"95de633a-083e-42f5-b444-a4295d8e9314" = "Microsoft Whiteboard Services"
    #"dfe74da8-9279-44ec-8fb2-2aed9e1c73d0" = "O365 SkypeSpaces Ingestion Service"
    #"4345a7b9-9a63-4910-a426-35363201d503" = "O365 Suite UX"
    #"00000002-0000-0ff1-ce00-000000000000" = "Office 365 Exchange Online"
    #"00b41c95-dab0-4487-9791-b9d2c32c80f2" = "Office 365 Management"
    #"66a88757-258c-4c72-893c-3e8bed4d6899" = "Office 365 Search Service"
    #"00000003-0000-0ff1-ce00-000000000000" = "Office 365 SharePoint Online"
    #"94c63fef-13a3-47bc-8074-75af8c65887a" = "Office Delve"
    #"93d53678-613d-4013-afc1-62e9e444a0a5" = "Office Online Add-in SSO"
    #"2abdc806-e091-4495-9b10-b04d93c3f040" = "Office Online Client Microsoft Entra ID- Augmentation Loop"
    #"b23dd4db-9142-4734-867f-3577f640ad0c" = "Office Online Client Microsoft Entra ID- Loki"
    #"17d5e35f-655b-4fb0-8ae6-86356e9a49f5" = "Office Online Client Microsoft Entra ID- Maker"
    #"b6e69c34-5f1f-4c34-8cdf-7fea120b8670" = "Office Online Client MSA- Loki"
    #"243c63a3-247d-41c5-9d83-7788c43f1c43" = "Office Online Core SSO"
    #"a9b49b65-0a12-430b-9540-c80b3332c127" = "Office Online Search"
    #"4b233688-031c-404b-9a80-a4f3f2351f90" = "Office.com"
    #"89bee1f7-5e6e-4d8a-9f3d-ecd601259da7" = "Office365 Shell WCSS-Client"
    #"0f698dd4-f011-4d23-a33e-b36416dcb1e6" = "OfficeClientService"
    #"4765445b-32c6-49b0-83e6-1d93765276ca" = "OfficeHome"
    #"4d5c2d63-cf83-4365-853c-925fd1a64357" = "OfficeShredderWacClient"
    #"62256cef-54c0-4cb4-bcac-4c67989bdc40" = "OMSOctopiPROD"
    #"ab9b8c07-8f02-4f72-87fa-80105867a763" = "OneDrive SyncEngine"
    #"2d4d3d8e-2be3-4bef-9f87-7875a61c29de" = "OneNote"
    #"27922004-5251-4030-b22d-91ecd9a37ea4" = "Outlook Mobile"
    #"a3475900-ccec-4a69-98f5-a65cd5dc5306" = "Partner Customer Delegated Admin Offline Processor"
    #"bdd48c81-3a58-4ea9-849c-ebea7f6b6360" = "Password Breach Authenticator"
    #"35d54a08-36c9-4847-9018-93934c62740c" = "PeoplePredictions"
    #"00000009-0000-0000-c000-000000000000" = "Power BI Service"
    #"ae8e128e-080f-4086-b0e3-4c19301ada69" = "Scheduling"
    #"ffcb16e8-f789-467c-8ce9-f826a080d987" = "SharedWithMe"
    #"08e18876-6177-487e-b8b5-cf950c1e598c" = "SharePoint Online Web Client Extensibility"
    #"b4bddae8-ab25-483e-8670-df09b9f1d0ea" = "Signup"
    #"00000004-0000-0ff1-ce00-000000000000" = "Skype for Business Online"
    #"61109738-7d2b-4a0b-9fe3-660b1ff83505" = "SpoolsProvisioning"
    #"91ca2ca5-3b3e-41dd-ab65-809fa3dffffa" = "Sticky Notes API"
    #"13937bba-652e-4c46-b222-3003f4d1ff97" = "Substrate Context Service"
    #"26abc9a8-24f0-4b11-8234-e86ede698878" = "SubstrateDirectoryEventProcessor"
    #"a970bac6-63fe-4ec5-8884-8536862c42d4" = "Substrate Search Settings Management Service"
    #"905fcf26-4eb7-48a0-9ff0-8dcc7194b5ba" = "Sway"
    #"97cb1f73-50df-47d1-8fb0-0271f2728514" = "Transcript Ingestion"
    #"268761a2-03f3-40df-8a8b-c3db24145b6b" = "Universal Store Native Client"
    #"00000005-0000-0ff1-ce00-000000000000" = "Viva Engage (formerly Yammer)"
    #"fe93bfe1-7947-460a-a5e0-7a5906b51360" = "Viva Insights"
    #"3c896ded-22c5-450f-91f6-3d1ef0848f6e" = "WeveEngine"
    #"00000002-0000-0000-c000-000000000000" = "Windows Azure Active Directory"
    #"8edd93e1-2103-40b4-bd70-6e34e586362d" = "Windows Azure Security Resource Provider"
    #"797f4846-ba00-4fd7-ba43-dac1f8f63013" = "Windows Azure Service Management API"
    #"a3b79187-70b2-4139-83f9-6016c58cd27b" = "WindowsDefenderATP Portal"
    #"26a7ee05-5602-4d76-a7ba-eae8b7b67941" = "Windows Search"
    #"1b3c667f-cde3-4090-b60b-3d2abd0117f0" = "Windows Spotlight"
    #"45a330b1-b1ec-4cc1-9161-9f03992aa49f" = "Windows Store for Business"
    #"c1c74fed-04c9-4704-80dc-9f79a2e515cb" = "Yammer Web"
    #"e1ef36fd-b883-4dbf-97f0-9ece4b576fc6" = "Yammer Web Embed"
    #"8ad40a1a-2a47-4760-ab30-c4360327d083" = "PnP Interactive"
}

# Define the start time for the query based on the DaysToQuery parameter
# Uses UTC time ("Z") as recommended for Graph API queries
$startTime = (Get-Date).AddDays(-$DaysToQuery).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

# Construct the full path for the Excel export file
$dateString = Get-Date -Format "ddMMyyHHmm"
$excelFileName = "PowerShellModuleLogin_$($dateString).xlsx"
$excelFilePath = Join-Path -Path $ExportPath -ChildPath $excelFileName

# --- Function Definition ---

# Initialize an array to store all sign-in logs
# Using a generic list for potentially better performance when adding many items
$combinedSignIns = [System.Collections.Generic.List[object]]::new()

# Define a function to query sign-in logs for a specific application ID
function Get-SignInsForApp {
    param (
        [Parameter(Mandatory = $true)]
        [string]$appId,

        [Parameter(Mandatory = $true)]
        [string]$appName,

        [Parameter(Mandatory = $true)]
        [string]$QueryStartTime
    )

    Write-Host "Querying sign-ins for App: '$appName' (ID: $appId) since $startTime"

    # Define the filter clause for the specific app ID and time range
    $filterQuery = "(resourceId eq '$appId' or appId eq '$appId') and createdDateTime ge $startTime and (signInEventTypes/any(t: t eq 'interactiveUser' or t eq 'servicePrincipal' or t eq 'managedIdentity'))"


    # Construct the URI for the audit logs endpoint with filters
    $uri = "https://graph.microsoft.com/beta/auditLogs/signIns?filter=$filterQuery"

    # Initialize an array to store the results
    $appSignIns = [System.Collections.Generic.List[object]]::new()
    $nextLink = $uri

    # Retrieve all sign-in logs matching the filter, handling pagination
    try {
        # Use Get-MgAuditLogSignIn for potential future compatibility and simplified pagination handling
        # However, Invoke-MgGraphRequest is needed here for the beta endpoint and complex filter.
        do {
            Write-Verbose "Executing Graph API Request: $nextLink"
            # Invoke the Graph request. Requires ConsistencyLevel header for advanced queries.
            $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -Headers @{ "ConsistencyLevel" = "eventual" } -ErrorAction Stop

            if ($null -ne $response.value) {
                $appSignIns.AddRange($response.value)
            }
            # Get the link for the next page of results, if any
            $nextLink = $response.'@odata.nextLink'
        } while ($null -ne $nextLink)
        
        # Append the appName to each sign-in log entry
        foreach ($signIn in $appSignIns) {
            $signIn | Add-Member -MemberType NoteProperty -Name "AppName" -Value $appName -Force # Use -Force to overwrite if property exists
        }

        Write-Host "Found $($appSignIns.Count) sign-ins for App ID: $AppId ('$AppName')"
        return $appSignIns
    }
    catch {
        # Log the error details
        Write-Error "Error retrieving sign-ins for App ID '$AppId' ('$AppName'). URI: $uri. Error: $($_.Exception.Message)"
        # Return an empty list to allow the script to continue with other apps
        return [System.Collections.Generic.List[object]]::new()
    }
}

# --- Main Script Logic ---

# Initialize a list to store all combined sign-in logs
$combinedSignIns = [System.Collections.Generic.List[object]]::new()

# Retrieve and combine sign-ins for each app ID
Write-Host "Starting sign-in log retrieval for configured applications..."
foreach ($appEntry in $applications.GetEnumerator()) {
    $appId = $appEntry.Key
    $appName = $appEntry.Value
    # Call the function to get sign-ins for the current app and add results to the combined list
    $combinedSignIns += Get-SignInsForApp -AppId $appId -AppName $appName -QueryStartTime $startTime
}
Write-Host "Finished retrieving logs. Total sign-ins found across all apps: $($combinedSignIns.Count)"

# Check if any sign-ins were found
if ($combinedSignIns.Count -eq 0) {
    Write-Host "No sign-ins found for the specified applications and time period."
    # Optionally exit or skip Excel export
    # exit 0
}
else {
    # Select relevant properties for the report
    # Using calculated properties for potentially nested or complex fields if needed in the future
    $reportData = $combinedSignIns | Select-Object -Property @(
        'AppName' # Added friendly name
        'appDisplayName'
        'userDisplayName'
        'userPrincipalName'
        'createdDateTime'
        'ipAddress'
        'clientAppUsed'
        'conditionalAccessStatus'
        'isInteractive'
        'resourceDisplayName'
        'resourceId'
        'appId' # Client App ID
        'userType'
        'tokenIssuerType'
        'clientCredentialType' # Useful for Service Principal sign-ins
    )


    # Export the data to an Excel file with a data sheet and a pivot table
    Write-Host "Exporting data to Excel file: $excelFilePath"
    try {
        # Ensure the ImportExcel module is available
        if (-not (Get-Command Export-Excel -ErrorAction SilentlyContinue)) {
            throw "The 'ImportExcel' module is required but not found. Please install it using 'Install-Module ImportExcel -Scope CurrentUser'."
        }


        $formattedStartTime = (Get-Date $startTime).ToString("dd-MM-yyyy")

        # Create a pivot table
        $pivotParams = @{
            PivotTableName = "SignIn Summary Since $formattedStartTime"
            PivotRows      = @("appDisplayName", "userDisplayName")
            PivotData      = @{"userDisplayName" = "count" }
        }


        # Add the pivot table to the Excel file
        $reportData  | Export-Excel -Path $excelFilePath -WorksheetName "SignIn Logs Since $formattedStartTime" -TableName "SignInLogsTable" -TableStyle Medium6 -IncludePivotTable @pivotParams -AutoSize
        


        Write-Host "Successfully exported sign-in logs to $excelFilePath"
    }
    catch {
        Write-Error "Failed to export data to Excel. Path: $excelFilePath. Error: $($_.Exception.Message)"
    }
}

# --- Script End ---
Write-Host "Script execution finished."