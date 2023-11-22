# Get-ComputerInfo.ps1

# Get the username of the currently logged in user
$username = [Environment]::UserName
Write-Output "Username: $username"

# Get information about the computer's operating system
$os = Get-WmiObject -Class Win32_OperatingSystem
Write-Output "OS: $($os.Caption), Service Pack: $($os.ServicePackMajorVersion).$($os.ServicePackMinorVersion)"

# Get information about the computer's CPU
$cpu = Get-WmiObject -Class Win32_Processor
Write-Output "CPU: $($cpu.Name), Cores: $($cpu.NumberOfCores), Threads: $($cpu.ThreadCount)"

# Get information about the computer's memory
$memory = Get-WmiObject -Class Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum
$totalMemoryGB = [Math]::Round($memory.Sum / 1GB, 2)
Write-Output "Total Memory: $totalMemoryGB GB"

# Get information about the computer's disk drives
$drives = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType=3"
foreach ($drive in $drives) {
    $freeSpaceGB = [Math]::Round($drive.FreeSpace / 1GB, 2)
    $totalSpaceGB = [Math]::Round($drive.Size / 1GB, 2)
    Write-Output "Drive $($drive.DeviceID): Free Space: $freeSpaceGB GB, Total Space: $totalSpaceGB GB"
}

# Get information about the computer's network adapters
$adapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPAddress -ne $null }
foreach ($adapter in $adapters) {
    Write-Output "Network Adapter $($adapter.Description): IP Address: $($adapter.IPAddress[0]), MAC Address: $($adapter.MACAddress)"
}