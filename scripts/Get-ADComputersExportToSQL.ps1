# Import the required module for Active Directory
Import-Module ActiveDirectory

# Get all computers from Active Directory
$computers = Get-ADComputer -Filter * -Property *

# Create an array to store the results
$results = @()

# Loop through each computer
foreach ($computer in $computers) {
    # Create a custom object with the computer's information
    $result = [PSCustomObject] @{
        Name = $computer.Name
        DNSHostName = $computer.DNSHostName
        Description = $computer.Description
        Enabled = $computer.Enabled
        OperatingSystem = $computer.OperatingSystem
        OperatingSystemServicePack = $computer.OperatingSystemServicePack
        OperatingSystemVersion = $computer.OperatingSystemVersion
        Location = $computer.Location
        UserAccountControl = $computer.UserAccountControl
        PasswordLastSet = $computer.PasswordLastSet
        WhenCreated = $computer.WhenCreated
        WhenChanged = $computer.WhenChanged
        LastLogonTimestampDT = $computer.LastLogonTimestampDT
        ManagedBy = $computer.ManagedBy
        Owner = $computer.Owner
        CanonicalName = $computer.CanonicalName
        DistinguishedName = $computer.DistinguishedName
        AdditionalColumn = $computer.AdditionalColumn
    }

    # Add the result to the results array
    $results += $result
}

# Convert the results to a comma-separated string
$data_string = $results | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Select-Object -Skip 1

# Output the data string to the console
Write-Output $data_string

# Check if the directory exists and create it if it doesn't
$directory = "C:\ITOA"
if (!(Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}

# Export the data to a CSV file
$results | Export-Csv -Path "$directory\Get-ADComputers.csv" -NoTypeInformation
