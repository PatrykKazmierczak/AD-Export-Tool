$PSEmailServer = "compinfainsa-com.mail.protection.outlook.com"         # mail server name
$smtpPort = 25
$emailfrom = "itoa@compinfainsa.com"   # doesn't have to be real email, just in the form name@domain.ext
$emailto = "pkazmierczak@compinfainsa.com"    

# email report file
Send-MailMessage -To $emailto -From "$emailfrom" -Subject "ITOA: Loading AD Users to SRV-AD-POLSKA\SQLEXPRESS01" -Attachments $logFile