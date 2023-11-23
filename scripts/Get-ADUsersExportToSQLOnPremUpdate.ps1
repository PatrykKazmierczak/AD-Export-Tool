<#"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
       Function:                  Get-ScriptDirectory
       Purpose:                   Gets the current directory for the script
       Last Update /by:    		  23/03/2023 Massimiliano Ferrazzi

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
	"Can"t find the ActiveDirectory module, please ensure it"s installed.`n"
	}
ELSE {Import-Module ActiveDirectory}

$CHECK2 = Get-Module SQLPS
IF(!$CHECK2) {
	Write-Host -ForegroundColor Red `
	"Can"t find the SQLPS module, please ensure it"s installed.`n"
	}
ELSE {Import-Module SQLPS}

##########################################
##### Fixed Variables ####################
##########################################
$logFile = "C:\scripts\AD2SQLUsers-log.txt"
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
$TABLE = "tbl.AllADUsers" #	[AD2SQL_Data_TEST].[dbo].[ADExport]

<#Table Create String ---- MIRAR AIXO, ACTUALIZAR AMB LA V 2#>

$CREATE = "CREATE TABLE $TABLE (SamAccountName varchar(150) not null PRIMARY KEY, `
	UserPrincipalName varchar(max),`
    DisplayName varchar(max),`
    givenName nvarchar (max),`
    Surname nvarchar (max),`
    Title nvarchar (max),`
    Description nvarchar(max), `
    Department nvarchar(max),`
    company nvarchar(max),`
    Office nvarchar(max),`
    Manager nvarchar(max),`
    ManagerUPN nvarchar(max),`
    StreetAddress nvarchar(max),`
    City nvarchar(max),`
    State nvarchar(max),`
    postalCode nvarchar(max),`
    CountryCode nvarchar(max),`
    Country nvarchar(max),`
    EmailAddress nvarchar(max),`
    OfficePhone nvarchar(max),`
    HomePhone nvarchar(max),`
    Mobile nvarchar(max),`
    FAX nvarchar(max),`
    EmployeeID nvarchar(max),`
    EmployeeNumber nvarchar(max),`
    HomeDirectory nvarchar(max),`
    HomeDrive nvarchar(max),`
    WhenCreated datetime,`
    WhenChanged datetime,`
    LastLogonDate datetime,`
    LastBadPasswordAttempt datetime,`
    PasswordLastSet datetime,`
    PasswordExpired nvarchar(max),`
    PasswordNeverExpires nvarchar(max),`
    Enabled nvarchar(max),`
    DistinguishedName nvarchar(max),`
    CanonicaName nvarchar(max))"

# The full-featured query
# Use the Get-ADUser cmdlet to retrieve all Active Directory users
$users = Get-ADUser -Filter * -Properties *

# Create an array to store the results
$results = @()

# Iterate over each user and retrieve the manager
foreach ($user in $users) {
    # Retrieve the DistinguishedName of the manager property from the user object
    $managerDN = $user.Manager

    # If the user has a manager, retrieve the manager object and add their name and userPrincipalName to the results array
    if ($managerDN) {
        $manager = Get-ADUser -Identity $managerDN
        $result = [PSCustomObject] @{
            UserName = $user.Name
            SamAccountName = $user.SamAccountName
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            Name = $User.GivenName
            Surname = $user.sn
            Title = $user.title
            Description = $user.description
            Department = $user.department
            Company = $user.company
            Office = $user.Office
            ManagerName = $manager.Name
            ManagerUPN = $manager.UserPrincipalName
            streetAddress = $user.streetAddress
            City = $user.City
            State = $user.State
            PostalCode = $user.PostalCode
            CountryCode =$user.countryCode
            Country = $user.Country
            EmailAddress = $user.EmailAddress
            OfficePhone = $user.OfficePhone
            HomePhone = $user.HomePhone
            Mobile = $user.Mobile
            FAX = $user.FAX
            EmployeeID = $user.EmployeeID
            EmployeeNumber = $user.EmployeeNumber
            HomeDirectory = $user.HomeDirectory
            HomeDrive = $user.HomeDrive
            whenCreated = $user.whenCreated
            whenChanged = $user.whenChanged
            LastBadPasswordAttempt = $user.LastBadPasswordAttempt
            LastLogonDate = $user.LastLogonDate
            PasswordLastSet = $user.PasswordLastSet
            PasswordExpired = $user.PasswordExpired
            PasswordNeverExpires = $user.PasswordNeverExpires
            Enabled = $user.Enabled
            DistinguishedName = $user.DistinguishedName
            CanonicalName = $user.CanonicalName
            AccountExpirationDate = $user.AccountExpirationDate          
        }
        $results += $result
    } else {
        $result = [PSCustomObject] @{
            UserName = $user.Name
            SamAccountName = $user.SamAccountName
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            Name = $User.GivenName
            Surname = $user.sn
            Title = $user.title
            Description = $user.description
            Department = $user.department
            Company = $user.company
            Office = $user.Office
            ManagerName = "No Manager"
            ManagerUPN = "No Manager"
            streetAddress = $user.streetAddress
            City = $user.City
            State = $user.State
            PostalCode = $user.PostalCode
            CountryCode =$user.countryCode
            Country = $user.Country
            EmailAddress = $user.EmailAddress
            OfficePhone = $user.OfficePhone
            HomePhone = $user.HomePhone
            Mobile = $user.Mobile
            FAX = $user.FAX
            EmployeeID = $user.EmployeeID
            EmployeeNumber = $user.EmployeeNumber
            HomeDirectory = $user.HomeDirectory
            HomeDrive = $user.HomeDrive
            whenCreated = $user.whenCreated
            whenChanged = $user.whenChanged
            LastBadPasswordAttempt = $user.LastBadPasswordAttempt
            LastLogonDate = $user.LastLogonDate
            PasswordLastSet = $user.PasswordLastSet
            PasswordExpired = $user.PasswordExpired
            PasswordNeverExpires = $user.PasswordNeverExpires
            Enabled = $user.Enabled
            DistinguishedName = $user.DistinguishedName
            CanonicalName = $user.CanonicalName
            AccountExpirationDate = $user.AccountExpirationDate
        }
        $results += $result
    }
}

# Check if the directory exists and create it if it doesn't
$directory = "C:\ITOA"
if (!(Test-Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}

# Export the data to a CSV file
$results | Export-Csv -Path "$directory\Get-ADUsers.csv" -NoTypeInformation


