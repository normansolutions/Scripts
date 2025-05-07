# Connect to Microsoft Graph with the required scopes
Connect-MgGraph -Scopes "Application.ReadWrite.All", "Group.Read.All"

# Define the Application ID
$appId = "fb78d390-0c51-40cd-8e17-fdbfab77341b"

# Retrieve the Service Principal using the Application ID
$sp = Get-MgServicePrincipal -Filter "appId eq '$appId'"
if (-not $sp) {
    # Create a new Service Principal if it doesn't exist
    New-MgServicePrincipal -AppId $appId
    # Retrieve the newly created Service Principal
    $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'"
}

# Retrieve the group with the display name 'PowerShell_Module_Allowed'
$group = Get-MgGroup -Filter "displayName eq 'PowerShell_Module_Allowed'"

# Define the properties to update the Service Principal
$ServicePrincipalUpdate = @{
    "accountEnabled" = "true"
    "appRoleAssignmentRequired" = "true"
}

# Update the Service Principal with the defined properties
Update-MgServicePrincipal -ServicePrincipalId $sp.Id -BodyParameter $ServicePrincipalUpdate

# Get existing role assignments for the Service Principal
$existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id
# Check if the role is already assigned
$roleAlreadyAssigned = $existingAssignments | Where-Object { $_.AppRoleId -eq $params.AppRoleId }

if ($roleAlreadyAssigned) {
    Write-Output "The role is already assigned to the service principal."
} else {
    # Proceed with the new role assignment
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $group.Id -ResourceId $sp.Id -Id ([Guid]::Empty.ToString())
}

# TESTING AND RESOLVING INCORRECT ASSIGNMENTS

# Using Graph API to list all app role assignments
$graphUrl = "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.Id)/appRoleAssignedTo"
$existingAssignments = Invoke-MgGraphRequest -Uri $graphUrl
$existingAssignments.value

# Remove a specific app role assignment if required (based on above results)
Remove-MgServicePrincipalAppRoleAssignment -AppRoleAssignmentId 'zL7bEHfpaEykKN0WfzqNVm0nmvr5h9RDukxeW4eOXT8' -ServicePrincipalId $sp.Id

# Retrieve the Service Principal again to check if app role assignment is required
$sp = Get-MgServicePrincipal -Filter "appId eq '$appId'"
$sp.appRoleAssignmentRequired

# Display the app role assignments of the Service Principal
$sp | Select-Object AppRoleAssignments
