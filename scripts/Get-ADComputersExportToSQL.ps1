<#"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
       Function:                Get-ScriptDirectory
       Purpose:                 Gets the current directory for the script
       Last Update /by:    		23/03/2023 Massimiliano Ferrazzi
                                         PRD
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""#>
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

########################################
$CHECK = Get-Module ActiveDirectory
IF(!$CHECK) {
	Write-Host -ForegroundColor Red `
	"Can"t find the ActiveDirectory module, please insure it"s installed.`n"
	}
ELSE {Import-Module ActiveDirectory}

$CHECK2 = Get-Module SQLPS
IF(!$CHECK2) {
	Write-Host -ForegroundColor Red `
	"Can"t find the SQLPS module, please insure it"s installed.`n"
	}
ELSE {Import-Module SQLPS}

##########################################
##### Fixed Variables ####################
##########################################
$logFile = "C:\scripts\AD2SQLComputers-log.txt"
$ScriptPath = Get-ScriptDirectory
# Generic variables
$Date = get-date -f dd-MM-yyyy
$Hour = get-date -f hh:mm:ss

$log = "==============================================================="  > $logFile
$Log = $env:USERNAME + " has Started the script at $Date $Hour" >> $logFile
$log = "==============================================================="  >> $logFile
$log = "Script Location = $ScriptPath at $env:COMPUTERNAME"  >> $logFile
$log = "==============================================================="  >> $logFile
<#DataBase Name:#>
$DB = "ITOA"

<#SQL Server Name:#>
$SQLSRVR = "SRV-AD-POLSKA\SQLEXPRESS01"

<#Table to Create:#>
$TABLE = "tbl.AllADComputers" #	[AD2SQL_Data_TEST].[dbo].[ADExport]

<#Table Create String ---- MIRAR AIXO, ACTUALIZAR AMB LA V 2#>

$CREATE = "CREATE TABLE $TABLE (ComputerName varchar(150) not null PRIMARY KEY,`
	FQDN nvarchar(max),`
	Description nvarchar(max),`
	Enabled varchar(max),`
	OS nvarchar(max),`
	ServicePack nvarchar(max),`
	OS_Version varchar(max),`
	userAccountControl nvarchar(max),`
	PasswordLastSet datetime,`
	whenCreated datetime,`
	whenChanged datetime,`
	LastLogonTimestampDT nvarchar(max),`
	Owner nvarchar(max),`
	CanonicalName nvarchar(max),`
	DistinguishedName nvarchar(max))"

# The full-featured query
# ... your existing script up to the point where you have the data ...

# The full-featured query
# ... your existing script up to the point where you have the data ...

# The full-featured query
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

# Check if the directory exists and create it if it doesn't
$directory = "C:\ITOA"
if (!(Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}

# Export the data to a CSV file
$ADObjects | Export-Csv -Path "$directory\Get-ADComputers.csv" -NoTypeInformation

# ... your existing script after the point where you have the data ...
