# Create an array to store the results
$results = @()

# Get the username of the currently logged in user
$username = [Environment]::UserName

# Get information about the computer's operating system
$os = Get-WmiObject -Class Win32_OperatingSystem

# Get information about the computer's CPU
$cpu = Get-WmiObject -Class Win32_Processor

# Get information about the computer's memory
$memory = Get-WmiObject -Class Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
$totalMemoryGB = [Math]::Round($memory.Sum / 1GB, 2)

# Get information about the computer's disk drives
$drives = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"
$driveInfo = foreach ($drive in $drives) {
    $freeSpaceGB = [Math]::Round($drive.FreeSpace / 1GB, 2)
    $totalSpaceGB = [Math]::Round($drive.Size / 1GB, 2)
    "Drive $($drive.DeviceID): Free Space: $freeSpaceGB GB, Total Space: $totalSpaceGB GB"
}

# Create a custom object with the information
$result = [PSCustomObject] @{
    Username = $username
    OS = "$($os.Caption), Service Pack: $($os.ServicePackMajorVersion).$($os.ServicePackMinorVersion)"
    CPU = "$($cpu.Name), Cores: $($cpu.NumberOfCores), Threads: $($cpu.ThreadCount)"
    TotalMemoryGB = $totalMemoryGB
    DriveInfo = $driveInfo -join '; '
    AdapterInfo = $adapterInfo -join '; '
}

# Add the result to the results array
$results += $result

# Convert the results to a comma-separated string
$data_string = $results | Out-String

# Output the data string to the console
Write-Output $data_string