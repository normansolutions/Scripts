<#
.SYNOPSIS
    Retrieves Microsoft Planner tasks and their comments from all plans within a specified Microsoft 365 Group using Microsoft Graph API.

.DESCRIPTION
    This script connects to Microsoft Graph, retrieves all tasks from all Planner Plans within a specified Microsoft 365 Group,
    fetches comments for each task, and outputs the combined data in a structured JSON format.
    It's optimized for efficiency using caching, pagination, and improved error handling.

    Handles former members gracefully by identifying them as "Former Member".

.PARAMETER GroupId
    The ID of the Microsoft 365 Group to retrieve Planner data from.

.PARAMETER DelayMilliseconds
    The delay in milliseconds between API calls to avoid throttling. Defaults to 50. Increase this value if you encounter throttling errors.

.EXAMPLE
    .\GetPlannerTasksAndComments2.ps1 -GroupId "yourGroupIdHere"

.EXAMPLE
    .\GetPlannerTasksAndComments2.ps1 -GroupId "yourGroupIdHere" -DelayMilliseconds 250

.NOTES
    Requires the following PowerShell modules:
        - Microsoft.Graph.Authentication
        - Microsoft.Graph.Planner
        - Microsoft.Graph.Groups
    
    The user running this script must have appropriate permissions to access Planner data within the specified Microsoft 365 Group.
    The script is optimized for performance with large task sets using caching and pagination.
    Former members are handled gracefully and identified as "Former Member".
#>
param (
    [Parameter(Mandatory = $true)]
    [string]$GroupId,
    [Parameter(Mandatory = $false)]
    [int]$DelayMilliseconds = 50
)

# Clear the console and display the script start time.
Clear-Host
$startTime = Get-Date
Write-Host "Script started at: $startTime"

#region Helper Functions

# Cache for user information to avoid repeated lookups.
$script:userCache = @{}

# Function to get user details with caching.
function Get-CachedUserDetails {
    param (
        [string]$userId
    )

    # Check if user details are already in the cache.
    if (-not $script:userCache.ContainsKey($userId)) {
        try {
            # Retrieve user details from Microsoft Graph.
            $user = Get-MgUser -UserId $userId -ErrorAction Stop
            # Store user details in the cache.
            $script:userCache[$userId] = @{
                DisplayName = $user.DisplayName
                Id          = $user.Id
            }
        }
        catch {
            # Handle invalid user gracefully.
            if ($_.Exception.Message -like '*[Request_ResourceNotFound]*') {
                Write-Host "User $userId not found - using Former" -ForegroundColor Blue
                # Mark user as a former member in the cache.
                $script:userCache[$userId] = @{
                    DisplayName = "Former Member"
                    Id          = $userId
                }
            }
            else {
                Write-Warning "Could not get details for user $userId : $_"
                # Mark user as a former member in the cache if an error occurs.
                $script:userCache[$userId] = @{
                    DisplayName = "Former Member"
                    Id          = $userId
                }
            }
        }
    }
    # Return user details from the cache.
    return $script:userCache[$userId]
}

# Function to get all tasks for a given plan using pagination.
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
            # Introduce a delay to avoid throttling.
            Start-Sleep -Milliseconds $DelayMilliseconds

            # Retrieve tasks from Microsoft Graph.
            $response = Invoke-MgGraphRequest -Uri $uri -Method Get -Headers $headers -ErrorAction Stop

            # Check if the response is valid and contains tasks.
            if ($response -and $response.value) {
                # Add each task to the allTasks array.
                foreach ($task in $response.value) {
                    $allTasks += $task
                }
            }
            else {
                Write-Warning "No tasks found in response for plan: $PlanId"
                break # Exit the loop if no tasks are found.
            }

            # Check if there's a nextLink for pagination.
            $uri = $response.'@odata.nextLink'

        } while ($uri) # Continue while there is a next page.
    }
    catch {
        Write-Warning "Failed to retrieve tasks for plan: $PlanId - $_"
    }

    return $allTasks
}

# Function to sanitize comment content by removing HTML tags and extra text.
function Sanitize-Comment {
    param (
        [string]$InputString
    )

    # Remove all HTML tags except for <a> tags with href attributes.
    $cleanedString = [regex]::Replace($InputString, '<(?!a\s+href=\[).*?>', '')

    # Remove the "Reply in Microsoft Planner" phrase and everything after it.
    $position = $cleanedString.IndexOf("Reply in Microsoft Planner")
    if ($position -ge 0) {
        $updatedComment = $cleanedString.Substring(0, $position)
    }
    else {
        $updatedComment = $cleanedString
    }

    # Remove leading and trailing carriage returns and newlines.
    $trimmedLeadingString = $updatedComment.TrimStart("`r", "`n")
    $trimmedString = $trimmedLeadingString.TrimEnd("`n", "`r")
    $result = $trimmedString + "`n"

    return $result
}

# Function to process task assignees and cache their details.
function Process-TaskAssignees {
    param (
        [array]$tasks
    )

    # Collect all unique user IDs from task assignments.
    $uniqueUserIds = @{}
    foreach ($task in $tasks) {
        if ($task.Assignments) {
            $assignmentData = $task.Assignments | Select-Object
            foreach ($userId in $assignmentData.Keys) {
                $uniqueUserIds[$userId] = $true
            }
        }
    }

    # Retrieve and cache details for each unique user.
    $uniqueUserIds.Keys | ForEach-Object {
        Get-CachedUserDetails -userId $_
    }
}

# Function to get the Microsoft 365 Group ID for a given Planner Plan ID.
function Get-PlannerPlanGroupId {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PlanId
    )

    try {
        # Try v1.0 endpoint first.
        $v1Url = "https://graph.microsoft.com/v1.0/planner/plans/$PlanId"
        $planDetails = Invoke-MgGraphRequest -Method GET -Uri $v1Url -ErrorAction Stop

        # Check for container properties in the response.
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

        # Fallback to beta endpoint if v1.0 fails.
        Write-Warning "Falling back to beta endpoint for plan details..."
        $betaUrl = "https://graph.microsoft.com/beta/planner/plans/$PlanId"
        $planDetails = Invoke-MgGraphRequest -Method GET -Uri $betaUrl -ErrorAction Stop

        # Check for container properties in the response.
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

        # Check for owner properties in the response.
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
        Write-Warning "Failed to retrieve Group ID using direct API calls: $_"
        return $null
    }
}

# Function to get comments for a specific task.
function Get-TaskComments {
    param (
        [string]$groupId,
        [string]$conversationThreadId,
        [int]$DelayMilliseconds
    )

    $comments = @()

    # Only proceed if both GroupId and ConversationThreadId are provided.
    if ($groupId -and $conversationThreadId) {
        try {
            # Introduce a delay to avoid throttling.
            Start-Sleep -Milliseconds $DelayMilliseconds
            # Retrieve comments from Microsoft Graph.
            $taskComments = Get-MgGroupThreadPost -GroupId $groupId -ConversationThreadId $conversationThreadId -ErrorAction Stop
            if ($taskComments) {
                foreach ($comment in $taskComments) {
                    # Sanitize the comment content.
                    $commentContent = Sanitize-Comment -InputString $comment.Body.Content

                    # Create a custom object for each comment.
                    $comments += [PSCustomObject]@{
                        PSTypeName  = 'PlannerTaskComment'
                        CommentDate = $comment.createdDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ")
                        CommentUser = $comment.Sender.EmailAddress.Name
                        Comment     = $commentContent
                    }
                }
            }
        }
        catch {
            Write-Warning "Error getting comments for conversation thread $conversationThreadId : $_"
        }
    }
    return $comments
}

# Function to get all Planner Plans within a specified Microsoft 365 Group.
function Get-AllPlannerPlansInGroup {
    param (
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )

    $allPlans = @()
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/planner/plans"

    try {
        do {
            # Retrieve plans from Microsoft Graph.
            $response = Invoke-MgGraphRequest -Uri $uri -Method Get -ErrorAction Stop

            # Check if the response is valid and contains plans.
            if ($response -and $response.value) {
                # Add each plan to the allPlans array.
                $allPlans += $response.value
            }
            else {
                Write-Warning "No plans found in group: $GroupId"
                break # Exit the loop if no plans are found.
            }

            # Check if there's a nextLink for pagination.
            $uri = $response.'@odata.nextLink'
        } while ($uri) # Continue while there is a next page.
    }
    catch {
        Write-Warning "Failed to retrieve plans for group: $GroupId - $_"
    }

    return $allPlans
}

#endregion

#region Module Installation and Authentication

# Check if required modules are installed and install if necessary.
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.Planner", "Microsoft.Graph.Groups")

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -ErrorAction Stop
        }
        catch {
            Write-Error "Failed to install $module module. Error: $_"
            exit 1
        }
    }

    try {
        Import-Module $module -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to import $module module. Error: $_"
        exit 1
    }
}

# Connect to Microsoft Graph with required permissions.
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
    Connect-MgGraph -ErrorAction Stop
    $context = Get-MgContext
    if (-not $context) {
        throw "Failed to connect to Microsoft Graph"
    }
    Write-Host "Successfully connected to Microsoft Graph as $($context.Account)" -ForegroundColor Green
}
catch {
    Write-Error "Authentication failed: $_"
    exit 1
}

#endregion

#region Main Logic

# Initialize an array to store the combined results.
$allResults = @()

# Get all plans in the specified group.
Write-Host "Retrieving all plans in group: $GroupId" -ForegroundColor Yellow
$allPlans = Get-AllPlannerPlansInGroup -GroupId $GroupId

# Exit if no plans are found.
if ($allPlans.Count -eq 0) {
    Write-Warning "No Planner plans found in group: $GroupId. Exiting."
    exit
}

# Process each plan.
foreach ($plan in $allPlans) {
    $PlanId = $plan.id
    try {
        Write-Host "Processing Plan: $($plan.title) (ID: $PlanId)" -ForegroundColor Yellow

        # Retrieve all tasks for the current plan.
        Write-Host "Retrieving tasks..." -ForegroundColor Yellow
        $tasks = Get-AllMgPlannerPlanTasks -PlanId $PlanId -DelayMilliseconds $DelayMilliseconds
        Write-Host "Retrieved $($tasks.Count) tasks" -ForegroundColor Green

        # Retrieve all buckets for the current plan.
        Write-Host "Retrieving buckets for plan $($plan.title)..." -ForegroundColor Yellow
        $buckets = @{}
        $planBuckets = Get-MgPlannerPlanBucket -PlannerPlanId $PlanId -ErrorAction SilentlyContinue
        if ($planBuckets) {
            foreach ($bucket in $planBuckets) {
                $buckets[$bucket.Id] = $bucket
            }
        }
        Write-Host "Retrieved $($buckets.Count) buckets for plan $($plan.title)" -ForegroundColor Green

        # Process task assignees and cache their details.
        if ($tasks.Count -gt 0) {
            Process-TaskAssignees -tasks $tasks
        }

        # Process tasks and comments sequentially.
        Write-Host "Processing tasks and retrieving comments for plan $($plan.title)..." -ForegroundColor Yellow
        $planResults = @()
        foreach ($task in $tasks) {
            # Get task details.
            $taskDetails = Get-MgPlannerTaskDetail -PlannerTaskId $task.Id -ErrorAction SilentlyContinue

            # Process assignments using the cache.
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

            # Get Bucket Name.
            $bucketName = "Unknown Bucket"
            if ($task.BucketId -and $buckets.ContainsKey($task.BucketId)) {
                $bucketName = $buckets[$task.BucketId].Name
            }

            # Get comments for the current task.
            $comments = Get-TaskComments -groupId $GroupId -conversationThreadId $task.ConversationThreadId -DelayMilliseconds $DelayMilliseconds

            # Create a custom object for the task and its details.
            $planResults += [PSCustomObject]@{
                taskId          = $task.Id
                taskName        = $task.Title
                taskDescription = if ($taskDetails) { $taskDetails.Description } else { $null }
                taskProgress    = $task.PercentComplete
                taskStart       = $task.StartDateTime
                taskDue         = $task.DueDateTime
                taskComplete    = $task.CompletedDateTime
                assignees       = $assignees
                buckets         = @(
                    [PSCustomObject]@{
                        bucketId   = $task.BucketId
                        bucketName = $bucketName
                    }
                )
                planName        = $plan.title
                comments        = @($comments)
                taskCreated     = $task.createdDateTime
                taskPriority    = $task.priority
            }
        }

        # Add the plan's results to the combined results.
        $allResults += $planResults
        Write-Host "Finished processing plan $($plan.title)" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to process plan $PlanId $_"
    }
}

#endregion

#region Export Results

try {
    # Generate a timestamp for the output filename.
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $outputFileName = "PlannerTasksExport-$timestamp.json"

    # Check if there are any results to export.
    if ($allResults.Count -eq 0) {
        Write-Warning "No results to export."
    }
    else {
        # Export the results to a JSON file.
        Write-Host "Exporting $($allResults.Count) tasks to $outputFileName" -ForegroundColor Yellow
        $allResults | ConvertTo-Json -Depth 20 -Compress:$false -EscapeHandling EscapeNonAscii | Out-File -FilePath $outputFileName -Encoding UTF8
        Write-Host "Successfully exported results to $outputFileName" -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to export results to JSON: $_"
}
finally {
    # Disconnect from Microsoft Graph.
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Yellow
    }
    catch {
        Write-Warning "Failed to disconnect from Microsoft Graph: $_"
    }
}

#endregion

# Display the script end time.
$endTime = Get-Date
Write-Host "Script finished at: $endTime"
