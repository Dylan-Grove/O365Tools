$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName
. "$CreateEXOPSSession\CreateExoPSSession.ps1"

Connect-EXOPSSession

$emaildisable = Read-Host "Enter email address of User"
Get-MobileDevice -Mailbox $emaildisable | Remove-MobileDevice -confirm:$false
Set-CASMailbox -Identity $emaildisable -ActiveSyncEnabled $False 
