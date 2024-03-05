This will allow you to access a client's O365 with ISE

1. Close all powershell and ISE windows
2. Install the Exchange oline module (Install this first --- Microsoft.Online.CSE.PSModule.Client)
 - If this doesn't work. Log into any O365 portal, Navigate to Exchange admin Center > Hybrid. Click configure on "The Exchange Online Powershell Module"
3. Paste the script in your ISE or use it as a function anytime you need to use it. (Connect-ExchangeOnlineISE.ps1)