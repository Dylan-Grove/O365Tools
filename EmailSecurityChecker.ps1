# This script will pull DKIM, SPF, And DMARC for supplied email domains in contacts.txt


$List = Get-Content C:\users\dgrove\Desktop\contacts.txt
$Domains = @()
Foreach($Email in $List){
    If($Email -like "*@*"){ $Domain = $Email.Split("@")[1] }
    Else{ $Domain = $Email }

    If($Domain -like "*onmicrosoft*"){ Continue }
    Else{ $Domains += $Domain.ToLower() }
}

$Domains = $Domains | Sort | Select -Unique
$DomainList = @()

$ErrorActionPreference = "silentlycontinue"

Foreach($Domain in $Domains){
    Write-Host -ForegroundColor Gray "Checking $Domain"
    
    $DKIMSelector1 = (Resolve-DnsName -Type CNAME -Name "Selector1._domainkey.$Domain").NameHost
    $DKIMSelector2 = (Resolve-DnsName -Type CNAME -Name "Selector2._domainkey.$Domain").NameHost
    $SPF           = ((Resolve-DnsName -Type TXT   -Name "$Domain" | Where-Object {$_.strings -like "*v=spf1*"}).Strings)
    $DMARC         = (Resolve-DnsName -Type TXT   -Name "_dmarc.$Domain").Strings


    $DomainList += (
        [pscustomobject]@{
            Domain=$Domain
            DKIMSelector1=$DKIMSelector1
            DKIMSelector2=$DKIMSelector2
            SPF=$SPF
            DMARC=$DMARC
        }
    )
}

$DomainList | select Domain,DkimSelector1,DKIMSelector2,@{Name=’SPF’;Expression={[string]::join(“;”, ($_.SPF))}},@{Name=’DMARC’;Expression={[string]::join(“;”, ($_.DMARC))}} | export-csv C:\users\dgrove\Desktop\domainlist.csv