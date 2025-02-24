# Function to generate a new GUID
function New-Guid {
    [System.Guid]::NewGuid().ToString()
}

# Function to create JSON entries
function Create-JsonEntries {
    param (
        [DateTime]$StartDate,       # Start date for the entries
        [DateTime]$EndDate,         # End date for the entries
        [string]$JsonFilePath,      # Path to the JSON file
        [string]$Description        # Description for each entry
    )

    # Check if the JSON file exists
    if (-Not (Test-Path -Path $JsonFilePath)) {
        Write-Error "JSON file not found: $JsonFilePath"
        return
    }

    # Read the existing JavaScript file
    $jsContent = Get-Content -Path $JsonFilePath -Raw

    # Remove 'var gData =' from the JavaScript file
    $jsContent = $jsContent -replace 'var gData =', ''

    # Convert the remaining content to JSON
    $json = $jsContent | ConvertFrom-Json

    # Debugging output to check the JSON structure
    Write-Output "JSON structure before modification:"
    $json | ConvertTo-Json -Depth 3

    # Check if GDates property exists, if not, initialize it
    if (-not $json.PSObject.Properties['GDates']) {
        Write-Output "GDates property not found. Initializing it."
        $json | Add-Member -MemberType NoteProperty -Name GDates -Value @()
    } else {
        Write-Output "GDates property found."
    }

    # Loop through each day between the start and end dates
    for ($date = $StartDate; $date -le $EndDate; $date = $date.AddDays(1)) {
        $entry = @{
            Id = New-Guid
            Date = $date.ToString("ddd dd MMM yyyy") + " (" + $date.ToString("dd/MM/yyyy") + ")"
            Description = $Description
        }
        $json.GDates += $entry
    }

    # Debugging output to check the JSON structure after modification
    Write-Output "JSON structure after modification:"
    $json | ConvertTo-Json -Depth 3

    # Convert back to JSON and re-add 'var gData =' at the beginning
    $newJsContent = "var gData =" + ($json | ConvertTo-Json -Depth 3)

    # Save the modified content back to the JavaScript file
    Set-Content -Path $JsonFilePath -Value $newJsContent
}

# Example usage
$startDate = [DateTime]::Parse("06-11-2025")  # Define the start date
$endDate = [DateTime]::Parse("06-12-2025")    # Define the end date
$jsonFilePath = "$PSScriptRoot\data\gdata.js"  # Path to the JSON file
$description = "Data Entry Date"  # Description for each entry

# Call the function to create JSON entries
Create-JsonEntries -StartDate $startDate -EndDate $endDate -JsonFilePath $jsonFilePath -Description $description