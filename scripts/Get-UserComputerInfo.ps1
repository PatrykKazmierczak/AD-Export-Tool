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

# Get information about the computer's network adapters
$adapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null }
$adapterInfo = foreach ($adapter in $adapters) {
    "Network Adapter $($adapter.Description): IP Address: $($adapter.IPAddress[0]), MAC Address: $($adapter.MACAddress)"
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

# Export the results to a CSV file
$results | Export-Csv -Path "C:\ITOA\ComputerInfo.csv" -NoTypeInformation