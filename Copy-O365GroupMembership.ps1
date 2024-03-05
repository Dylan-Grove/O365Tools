Function Copy-O365GroupMembership{
    
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true)][String]$SourceGroup,
        [parameter(Mandatory=$true)][String]$DestinationGroup
    )

    If($NewSession = "Basic"){ New-O365Session -Basic }
    If($NewSession = "Modern"){ New-O365Session -Modern }

    Get-DistributionGroupMember -Identity $SourceGroup | % {Write-Host -ForegroundColor Yellow "Attempting to add" $_.Name "..." ;Add-DistributionGroupMember -Identity $DestinationGroup -Member $_.PrimarySmtpAddress}

}
