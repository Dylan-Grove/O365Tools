# Copy the Source content of an email and run this command to pull the sender IP Address
# Good for whitelisting

$Clipboard = Get-Clipboard

Select-String -InputObject $Clipboard -Pattern '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b' | % {$_.Matches.Value}
