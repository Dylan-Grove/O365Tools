# Digest CSV
$csv = import-csv "Groups.csv"
$Users = @()

Foreach($i in $csv){
    If($i -eq ""){ Continue }
        
    $Last = ($i.Name).Split(" ")[1]
    $Last = $Last.Trim()
    $First = ($i.Name).Split(" ")[0]
    Try{$First = $First.Trim()}
    Catch{Continue}
    $Email = $First+$Last[0]+"@EMAIL.CA"

    $Users+=(
        [pscustomobject]@{
            First=$First
            Last=$Last
            Email=$Email
            Group1=$i.'Group 1'

        }
    )

}


$Users | ForEach-Object { 
    
    $Memberships = Get-Member -InputObject $_ -MemberType NoteProperty | select name,definition
    $DisplayName = ($_.First+" "+$_.Last)
    $Email = $_.Email
    $Groups = @()



    # Add to Groups
    Foreach($Group in $Memberships){
        If($Group.name -eq "First" -or $Group.name -eq "Last" -or $Group.name -eq "Email"){
            Continue
        }
        ElseIf($Group.Definition -like "*TRUE*"){
            $Groups += $Group.Name
        }
    }
    
    
    Foreach ($Group in $Groups){
    $Check = ((Get-DistributionGroupMember -Identity "$Group*" | Where-Object { $_.displayname -like "*$DisplayName*" }))
        If(! $Check.DisplayName -eq $DisplayName){ 
            Add-DistributionGroupMember -Identity (Get-DistributionGroup "$Group*").id -Member $Email
            Write-Host -ForegroundColor Yellow "Added $DisplayName to $Group"
        }
        Else { Write-Host -ForegroundColor Gray ("$DisplayName already in group: $Group") }
        }

}

