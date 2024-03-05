[System.Net.ServicePointManager]::SecurityProtocol = 'TLS12'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$Date = (Get-Date -Format dd-MM-yyyy)
$EmailFrom = ""
$EmailTo = ""
$SMTPServer = ""

$Subject = "[Alert]"
$Body = "
"

Send-MailMessage -SmtpServer $SMTPServer -Credential (Get-Credential f12admin@gasalberta.com) -From $EmailFrom -To $EmailTo -Subject $Subject -UseSsl -Body $Body