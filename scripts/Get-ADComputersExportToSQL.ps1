Function Get-ScriptDirectory
{
       $Invocation = (Get-Variable MyInvocation -Scope 1).Value
       [string]$path = Split-Path $Invocation.MyCommand.Path
       if (! $path.EndsWith("\"))
       {
              $path += "\"
       }
       return $path
}

$CHECK = Get-Module ActiveDirectory
IF(!$CHECK) {
    Write-Host -ForegroundColor Red `
    "Can't find the ActiveDirectory module, please insure it's installed.`n"
    }
ELSE {Import-Module ActiveDirectory}

$CHECK2 = Get-Module SQLPS
IF(!$CHECK2) {
    Write-Host -ForegroundColor Red `
    "Can't find the SQLPS module, please insure it's installed.`n"
    }
ELSE {Import-Module SQLPS}

$logFile = "C:\scripts\AD2SQLComputers-log.txt"
$ScriptPath = Get-ScriptDirectory
$Date = get-date -f dd-MM-yyyy
$Hour = get-date -f hh:mm:ss

$log = "==============================================================="  > $logFile
$Log = $env:USERNAME + " has Started the script at $Date $Hour" >> $logFile
$log = "==============================================================="  >> $logFile
$log = "Script Location = $ScriptPath at $env:COMPUTERNAME"  >> $logFile
$log = "==============================================================="  >> $logFile

$DB = "ITOA"
$SQLSRVR = "SRV-AD-POLSKA\SQLEXPRESS01"
$TABLE = "tbl.AllADComputers"

$ADObjects = Get-ADComputer -Filter * -Properties * |
    Select-Object Name, DNSHostName, Description, Enabled, OperatingSystem, `
        OperatingSystemServicePack, OperatingSystemVersion, Location, userAccountControl, PasswordLastSet, `
        whenCreated, whenChanged, `
        @{name="LastLogonTimestampDT";`
            Expression={[datetime]::FromFileTimeUTC($_.LastLogonTimestamp)}}, 
            ManagedBy,`
        @{name="Owner";`
            Expression={$_.nTSecurityDescriptor.Owner}}, `
        CanonicalName, DistinguishedName

$directory = "C:\ITOA"
if (!(Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}

$ADObjects | Export-Csv -Path "$directory\Get-ADComputers.csv" -NoTypeInformation