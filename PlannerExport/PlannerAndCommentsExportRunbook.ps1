<#
.SYNOPSIS
    Retrieves Microsoft Planner tasks and their comments from all plans within a specified Microsoft 365 Group using Microsoft Graph API.
    This version is designed to run as an Azure Automation Runbook.

.DESCRIPTION
    This script connects to Microsoft Graph, retrieves all tasks from all Planner Plans within a specified Microsoft 365 Group,
    fetches comments for each task, and outputs the combined data in a structured format.
    This version is optimized for efficiency by using caching and improved error handling.

    Handles former members gracefully.

.PARAMETER GroupId
    The ID of the Microsoft 365 Group to retrieve Planner data from.

.PARAMETER DelayMilliseconds
    The delay in milliseconds between API calls to avoid throttling. Defaults to 50.

.EXAMPLE
    .\GetPlannerTasksAndComments2.ps1 -GroupId "yourGroupIdHere"

.EXAMPLE
    .\GetPlannerTasksAndComments2.ps1 -GroupId "yourGroupIdHere" -DelayMilliseconds 250

.NOTES
    Requires Microsoft.Graph.Authentication, Microsoft.Graph.Planner, and Microsoft.Graph.Groups modules.
    User must have appropriate permissions to access Planner data.
    Optimized for performance with large task sets using caching.
    Handles former members gracefully.
#>
$VerbosePreference = 'SilentlyContinue'
#region Parameters - Azure Runbook Specific
# Define variables for the runbook
$GroupId = Get-AutomationVariable -Name 'GroupID'
[int]$DelayMilliseconds = 50
#endregion

#region Logging Function - Azure Runbook Specific
# Function to write to the runbook output stream
function Write-RunbookLog {
    param (
        [string]$Message,
        [string]$Type = "Information"
    )

    switch ($Type) {
        "Information" { Write-Output $Message }
        "Warning" { Write-Warning -Message $Message }
        "Error" { Write-Error -Message $Message }
        default { Write-Output $Message }
    }
}

#endregion

# Display the start time
$startTime = Get-Date
Write-RunbookLog -Message "Script started at: $startTime"

#region Helper Functions

# Cache for user information to avoid repeated lookups
$script:userCache = @{}

# Function to get user details with caching
function Get-CachedUserDetails {
    param (
        [string]$userId
    )

    if (-not $script:userCache.ContainsKey($userId)) {
        try {
            $user = Get-MgUser -UserId $userId -ErrorAction Stop
            $script:userCache[$userId] = @{
                DisplayName = $user.DisplayName
                Id          = $user.Id
            }
        }
        catch {
            # Handle invalid user gracefully
            if ($_.Exception.Message -like '*[Request_ResourceNotFound]*') {
                Write-RunbookLog -Message "User $userId not found - using Former" -Type "Information"
                $script:userCache[$userId] = @{
                    DisplayName = "Former Member"
                    Id          = $userId
                }
            }
            else {
                Write-RunbookLog -Message "Could not get details for user $userId : $_" -Type "Warning"
                $script:userCache[$userId] = @{
                    DisplayName = "Former Member"
                    Id          = $userId
                }
            }
        }
    }
    return $script:userCache[$userId]
}


# Function to get all tasks using pagination
function Get-AllMgPlannerPlanTasks {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PlanId,
        [int]$DelayMilliseconds
    )    
    
    $allTasks = @()
    $uri = "https://graph.microsoft.com/v1.0/planner/plans/$PlanId/tasks"
    $headers = @{
        "ConsistencyLevel" = "eventual"
    }

    try {
        do {
            # Introduce a delay to avoid throttling
            Start-Sleep -Milliseconds $DelayMilliseconds

            $response = Invoke-MgGraphRequest -Uri $uri -Method Get -Headers $headers -ErrorAction Stop

            # Check if the response is valid and contains a value array
            if ($response -and $response.value) {
                foreach ($task in $response.value) {
                    $allTasks += $task
                }
            }
            else {
                Write-RunbookLog -Message "No tasks found in response for plan: $PlanId" -Type "Warning"
                break # Exit the loop if no tasks are found
            }

            # Check if there's a nextLink for pagination
            $uri = $response.'@odata.nextLink'

        } while ($uri)
    }
    catch {
        Write-RunbookLog -Message "Failed to retrieve tasks for plan: $PlanId - $_" -Type "Warning"
        # Consider adding more robust error handling here, like retries
    }

    return $allTasks
}

# Function to remove the last N lines from a string
function Sanitize-Comment {
    param (
        [string]$InputString
    )

    # Remove all HTML tags
    $cleanedString = [regex]::Replace($InputString, '<(?!a\s+href=\[).*?>', '')

    $position = $cleanedString.IndexOf("Reply in Microsoft Planner")

    # If the phrase is found, remove it and everything after it
    if ($position -ge 0) {
        $updatedComment = $cleanedString.Substring(0, $position)
    }
    else {
        $updatedComment = $cleanedString
    }

    # Remove leading carriage returns and newlines
    $trimmedLeadingString = $updatedComment.TrimStart("`r", "`n")
    $trimmedString = $trimmedLeadingString.TrimEnd("`n", "`r")
    $result = $trimmedString + "`n"

    return $result
}

# Batch process users for a set of tasks
function Process-TaskAssignees {
    param (
        [array]$tasks
    )

    # Collect all unique user IDs first
    $uniqueUserIds = @{}
    foreach ($task in $tasks) {
        if ($task.Assignments) {
            $assignmentData = $task.Assignments | Select-Object
            foreach ($userId in $assignmentData.Keys) {
                $uniqueUserIds[$userId] = $true
            }
        }
    }

    $uniqueUserIds.Keys | ForEach-Object {
        Get-CachedUserDetails -userId $_
    }
}

# Function to get Planner Plan Group ID using Microsoft Graph direct API call
function Get-PlannerPlanGroupId {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PlanId
    )

    try {
        # Try v1.0 first
        $v1Url = "https://graph.microsoft.com/v1.0/planner/plans/$PlanId"
        $planDetails = Invoke-MgGraphRequest -Method GET -Uri $v1Url -ErrorAction Stop

        if ($planDetails.container) {
            if ($planDetails.container.groupId) {
                return $planDetails.container.groupId
            }
            elseif ($planDetails.container.containerId -and $planDetails.container.type -eq 'group') {
                return $planDetails.container.containerId
            }
            elseif ($planDetails.container.url -match '/groups/([^/]+)/') {
                return $Matches[1]
            }
        }

        # Fallback to beta only if v1.0 fails
        Write-RunbookLog -Message "Falling back to beta endpoint for plan details..." -Type "Warning"
        $betaUrl = "https://graph.microsoft.com/beta/planner/plans/$PlanId"
        $planDetails = Invoke-MgGraphRequest -Method GET -Uri $betaUrl -ErrorAction Stop

        if ($planDetails.container) {
            if ($planDetails.container.groupId) {
                return $planDetails.container.groupId
            }
            elseif ($planDetails.container.containerId -and $planDetails.container.type -eq 'group') {
                return $planDetails.container.containerId
            }
            elseif ($planDetails.container.url -match '/groups/([^/]+)/') {
                return $Matches[1]
            }
        }

        if ($planDetails.owner) {
            foreach ($propName in @('groupId', 'group_id', 'GroupId', 'groupID')) {
                if ($planDetails.owner.$propName) {
                    return $planDetails.owner.$propName
                }
            }
        }

        return $null
    }
    catch {
        Write-RunbookLog -Message "Failed to retrieve Group ID using direct API calls: $_" -Type "Warning"
        return $null
    }
}

# Function to get comments directly (no batching)
function Get-TaskComments {
    param (
        [string]$groupId,
        [string]$conversationThreadId,
        [int]$DelayMilliseconds
    )

    $comments = @()

    if ($groupId -and $conversationThreadId) {
        try {
            Start-Sleep -Milliseconds $DelayMilliseconds
            $taskComments = Get-MgGroupThreadPost -GroupId $groupId -ConversationThreadId $conversationThreadId -ErrorAction Stop
            if ($taskComments) {
                foreach ($comment in $taskComments) {

                    # Clean up the comment content BEFORE removing HTML
                    $commentContent = Sanitize-Comment -InputString $comment.Body.Content

                    $comments += [PSCustomObject]@{
                        PSTypeName  = 'PlannerTaskComment'
                        CommentDate = $comment.createdDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
                        CommentUser = $comment.Sender.EmailAddress.Name
                        Comment     = $commentContent -replace '[“”"\/]', ''# Use the formatted comment
                    }
                }
            }
        }
        catch {
            Write-RunbookLog -Message "Error getting comments for conversation thread $conversationThreadId : $_" -Type "Warning"
        }
    }
    return $comments
}

# Function to get all plans in a group
function Get-AllPlannerPlansInGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    $allPlans = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/planner/plans"

    try {
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method Get -ErrorAction Stop

            if ($response -and $response.value) {
                $allPlans += $response.value
            }
            else {
                Write-RunbookLog -Message "No plans found in group: $GroupId" -Type "Warning"
                break
            }

            $uri = $response.'@odata.nextLink'
        } while ($uri)
    }
    catch {
        Write-RunbookLog -Message "Failed to retrieve plans for group: $GroupId - $_" -Type "Warning"
    }

    return $allPlans
}

#endregion


# Connect to Microsoft Graph using the Managed Identity
Write-RunbookLog -Message "Connecting to Microsoft Graph..."
try {
    # Connect using managed identity
    Connect-MgGraph -Identity -ErrorAction Stop
    $context = Get-MgContext
    if (-not $context) {
        throw "Failed to connect to Microsoft Graph"
    }
    Write-RunbookLog -Message "Successfully connected to Microsoft Graph as $($context.Account)"
}
catch {
    Write-RunbookLog -Message "Authentication failed: $_" -Type "Error"
    throw $_
}

#endregion

#region Main Logic

# Initialize combined results array
$allResults = @()

# Get all plans in the specified group
Write-RunbookLog -Message "Retrieving all plans in group: $GroupId"
$allPlans = Get-AllPlannerPlansInGroup -GroupId $GroupId

if ($allPlans.Count -eq 0) {
    Write-RunbookLog -Message "No Planner plans found in group: $GroupId. Exiting." -Type "Warning"
    exit
}

# Process each plan
foreach ($plan in $allPlans) {
    $PlanId = $plan.id
    try {
        Write-RunbookLog -Message "Processing Plan: $($plan.title) (ID: $PlanId)"

        Write-RunbookLog -Message "Retrieving tasks..."
        $tasks = Get-AllMgPlannerPlanTasks -PlanId $PlanId -DelayMilliseconds $DelayMilliseconds
        #$tasks = Get-AllMgPlannerPlanTasks -PlanId $PlanId -DelayMilliseconds $DelayMilliseconds | Where-Object { $_.Id -eq 'idtKzT5lQEOAnrmeDl9j05YAH8--' }
        Write-RunbookLog -Message "Retrieved $($tasks.Count) tasks"

        # Get all buckets for the current plan
        Write-RunbookLog -Message "Retrieving buckets for plan $($plan.title)..."
        $buckets = @{}
        $planBuckets = Get-MgPlannerPlanBucket -PlannerPlanId $PlanId -ErrorAction SilentlyContinue
        if ($planBuckets) {
            foreach ($bucket in $planBuckets) {
                $buckets[$bucket.Id] = $bucket
            }
        }
        Write-RunbookLog -Message "Retrieved $($buckets.Count) buckets for plan $($plan.title)"

        if ($tasks.Count -gt 0) {
            Process-TaskAssignees -tasks $tasks
        }

        # Process tasks and comments sequentially
        Write-RunbookLog -Message "Processing tasks and retrieving comments for plan $($plan.title)..."
        $planResults = @()
        foreach ($task in $tasks) {
            # Get task details
            $taskDetails = Get-MgPlannerTaskDetail -PlannerTaskId $task.Id -ErrorAction SilentlyContinue

            # Process assignments using cache
            $assignees = @()
            if ($task.Assignments) {
                $assignmentData = $task.Assignments | Select-Object
                foreach ($userId in $assignmentData.Keys) {
                    $userDetails = Get-CachedUserDetails -userId $userId
                    if ($userDetails) {
                        $assignees += [PSCustomObject]@{
                            assigneeId   = $userId
                            assigneeName = $userDetails.DisplayName
                        }
                    }
                }
            }

            # Get Bucket Name
            $bucketName = "Unknown Bucket"
            if ($task.BucketId -and $buckets.ContainsKey($task.BucketId)) {
                $bucketName = $buckets[$task.BucketId].Name
            }

            # Get comments directly (no batching)
            # Pass the task title to Get-TaskComments
            $comments = Get-TaskComments -groupId $GroupId -conversationThreadId $task.ConversationThreadId -DelayMilliseconds $DelayMilliseconds
            $planResults += [PSCustomObject]@{
                taskId          = $task.Id
                taskName        = $task.Title -replace '[“”"\/]', ''
                taskDescription = $taskDescription = if ($taskDetails) { $taskDetails.Description -replace '[“”"\/]', '' } else { $null }             
                taskProgress    = $task.PercentComplete
                taskStart       = if ($task.StartDateTime -ne $null) { $task.StartDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ") } else { $null }
                taskDue         = if ($task.DueDateTime -ne $null) { $task.DueDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ") } else { $null }
                taskComplete    = if ($task.CompletedDateTime -ne $null) { $task.CompletedDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ") } else { $null }
                assignees       = $assignees
                buckets         = @(
                    [PSCustomObject]@{
                        bucketId   = $task.BucketId
                        bucketName = $bucketName
                    }
                )
                planName        = $plan.title
                comments        = @($comments)  
                taskCreated     = if ($task.createdDateTime -ne $null) { $task.createdDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ") } else { $null }
                taskPriority    = $task.priority                
            }
        }

        $allResults += $planResults
        Write-RunbookLog -Message "Finished processing plan $($plan.title)"
    }
    catch {
        Write-RunbookLog -Message "Failed to process plan $PlanId $_" -Type "Error"
    }
}

#endregion

#region Export Results - Azure Runbook Specific
# Function to get access token using managed identity
function Get-AccessToken {
    param (
        [string]$resource
    )
    $tokenAuthUri = $env:IDENTITY_ENDPOINT + "?resource=$resource&api-version=2019-08-01"
    try {
        $tokenResponse = Invoke-RestMethod -Method GET -Headers @{"X-IDENTITY-HEADER"=$env:IDENTITY_HEADER} -Uri $tokenAuthUri
        return $tokenResponse.access_token
    }
    catch {
        Write-RunbookLog -Message "Failed to get access token: $_" -Type "Error"
        throw $_
    }
}

# Function to get site ID
function Get-SiteId {
    param (
        [string]$tenantName,
        [string]$siteName,
        [string]$accessToken
    )
    $siteIdUrl = "https://graph.microsoft.com/v1.0/sites/$($tenantName).sharepoint.com:/sites/$($siteName)"
    try {
        $siteResponse = Invoke-RestMethod -Uri $siteIdUrl -Headers @{Authorization = "Bearer $accessToken"}
        return $siteResponse.id
    }
    catch {
        Write-RunbookLog -Message "Failed to get site ID: $_" -Type "Error"
        throw $_
    }
}

# Function to get drive ID
function Get-DriveId {
    param (
        [string]$siteId,
        [string]$folderPath,
        [string]$accessToken
    )
    $drivesUrl = "https://graph.microsoft.com/v1.0/sites/$($siteId)/drives"
    try {
        $drivesResponse = Invoke-RestMethod -Uri $drivesUrl -Headers @{Authorization = "Bearer $accessToken"}
        $driveId = ($drivesResponse.value | Where-Object { $_.name -eq $folderPath }).id
        return $driveId
    }
    catch {
        Write-RunbookLog -Message "Failed to get drive ID: $_" -Type "Error"
        throw $_
    }
}

# Function to upload file
function Upload-File {
    param (
        [string]$driveId,
        [string]$fileName,
        [string]$jsonExport,
        [string]$accessToken
    )
    $uploadUrl = "https://graph.microsoft.com/v1.0/drives/$($driveId)/root:/$($fileName):/content"
    try {
        $uploadResponse = Invoke-RestMethod -Method Put -Uri $uploadUrl -Headers @{Authorization = "Bearer $accessToken"} -Body $jsonExport -ContentType "text/plain"
        Write-RunbookLog -Message "File uploaded successfully: $uploadResponse"
    }
    catch {
        Write-RunbookLog -Message "Failed to upload file: $_" -Type "Error"
        throw $_
    }
}

# Main script execution
try {
    $resource = "https://graph.microsoft.com"
    $accessToken = Get-AccessToken -resource $resource

    $tenantName = Get-AutomationVariable -Name 'TenantName'
    $siteName = Get-AutomationVariable -Name 'SiteName'
    $fileName = Get-AutomationVariable -Name 'FileName'
    $folderPath = Get-AutomationVariable -Name 'FolderPath'
    $jsonExport = $allResults | ConvertTo-Json -Depth 20


    $siteId = Get-SiteId -tenantName $tenantName -siteName $siteName -accessToken $accessToken
    $driveId = Get-DriveId -siteId $siteId -folderPath $folderPath -accessToken $accessToken

    Upload-File -driveId $driveId -fileName $fileName -jsonExport $jsonExport -accessToken $accessToken
}
catch {
    Write-RunbookLog -Message "Script execution failed: $_" -Type "Error"
}
finally {
    Write-RunbookLog -Message "Script execution completed"
}
#endregion

# Display the end time
$endTime = Get-Date
Write-RunbookLog -Message "Script finished at: $endTime"
