# Initialize the cache as a hash table
$script:userCache = @{}

function Get-CachedUserDetails {
    param (
        [string]$userId
    )
    # Capture the start time
    $startTime = Get-Date
    # Check if user details are already in the cache.
    if (-not $script:userCache.ContainsKey($userId)) {
        try {
            # Retrieve user details from Microsoft Graph.
            Write-Host "$userId not in cache - having to make call"
            $user = Get-MgUser -UserId $userId -ErrorAction Stop
            if ($null -ne $user) {
                # Store user details in the cache.
                $script:userCache[$userId] = @{
                    DisplayName = $user.DisplayName
                    Id          = $user.Id
                    UPN         = $user.UserPrincipalName
                }
            }
            else {
                Write-Host "No user details found for $userId."
            }
        }
        catch {
            # Handle invalid user gracefully.
            if ($_.Exception.Message -like '*[Request_ResourceNotFound]*') {
                Write-Host "User $userId not found." -ForegroundColor Blue
                # Do not add anything to the cache if the user is not found.
            }
            else {
                Write-Warning "Could not get details for user $userId : $_"
                # Do not add anything to the cache if an error occurs.
            }
        }
    }
    else {
        Write-Host "$userId already in cache"
    }
    # Return user details from the cache.
    # Capture the end time
    $endTime = Get-Date
    # Calculate the duration
    $duration = $endTime - $startTime 
    # Output the duration
    Write-Host "Execution Time: $($duration.TotalSeconds) seconds"
       
    return $script:userCache[$userId]

}


Connect-MgGraph

$Result = Get-CachedUserDetails -userId "testuser@testdomain.co.uk"
$Result2 = Get-CachedUserDetails -userId "testuser@testdomain.co.uk"
$Result3 = Get-CachedUserDetails -userId "testuser2@testdomain.co.uk"
$Result4 = Get-CachedUserDetails -userId "testuser2@testdomain.co.uk"