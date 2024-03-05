#Import user list
$Users = Import-Csv -Path thing.csv

$O365Users = @()
# For each user in the list build above, Search for them in AD and O365
$Users | ForEach-Object { 
    $O365Search = ""
    $DisplayName = ($_.First+" "+$_.Last)
    $First = $_.first
    $Last = $_.last
    $Username = $_.username

    $O365Search = Get-Mailbox $DisplayName
    If($O365Search.length -lt 2){ $O365Search = Get-Mailbox *$Last}
    
    if($O365Search -is [system.array]){
        Write-Host -ForegroundColor Yellow "Multiple users detected for $DisplayName, Possible Matches:"
        $i = 1
        $O365Search | % {Write-Host $i,$_; $i++}
        $Selection = read-host "-----> Make a selection"
        If($Selection -ne ""){ $O365Search = $O365Search[$Selection] }
        Elseif($O365Search -is [system.array]){$O365Search=""}

        }

    $O365Users += $O365Search
}

