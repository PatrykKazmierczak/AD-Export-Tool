
# Get the path to the desktop
$desktopPath = [Environment]::GetFolderPath("Desktop")

# Create the full path for the CSV file
$csvPath = Join-Path -Path $desktopPath -ChildPath 'AzureADUsers.csv'

# Export users to a CSV file on the desktop
$users | Export-Csv -Path $csvPath -NoTypeInformation

try {
    # Import the module
    Import-Module AzureAD -WarningAction SilentlyContinue

    # Connect to your Azure AD
    $credential = Get-Credential
    Connect-AzureAD -Credential $credential

    # Get all users
    $users = Get-AzureADUser -All $true

    # Get the path to the desktop
    $desktopPath = [Environment]::GetFolderPath("Desktop")

    # Create the full path for the CSV file
    $csvPath = Join-Path -Path $desktopPath -ChildPath 'AzureADUsers.csv'

    # Export users to a CSV file on the desktop
    $users | Export-Csv -Path $csvPath -NoTypeInformation
} catch {
    Write-Host "Error: $_"
}