$ErrorActionPreference = "Stop"

# Path for Logs and CSV
$Date = Get-Date -Format yyyyMMdd-hhmm
$Path = "C:\MonitoringLogs\O365ADAudit"
$LogPath = "$Path\$Date-O365ADAudit-Log.txt"
If(!(Test-path $Path)){mkdir $Path | Out-Null}


# Logging Function
Function Write-Log {
    Param(
        [parameter(Mandatory=$true)][String]$LogMessage,
        
        [ValidateSet("Info","Debug","ActionRequired","Error")]
        [parameter(Mandatory=$true)][String]$LogType
    )

    If($LogType -eq "Info"){
        $LogMessage = (Get-Date -format "yy/MM/dd hh:mm:ss")+", INFO,    $LogMessage"
        Write-host $LogMessage -foregroundcolor Gray
        $LogMessage>> $LogPath
    }
    ElseIf($LogType -eq "Debug"){
        $LogMessage = (Get-Date -format "yy/MM/dd hh:mm:ss")+", DEBUG,   $LogMessage"
        Write-host $LogMessage -foregroundcolor Yellow
        $LogMessage>> $LogPath
    }
    ElseIf($LogType -eq "ActionRequired"){
        $LogMessage = (Get-Date -format "yy/MM/dd hh:mm:ss")+", ACT REQ, $LogMessage"
        Write-host $LogMessage -foregroundcolor Cyan
        $LogMessage>> $LogPath
    }
    ElseIf($LogType -eq "Error"){
        $LogMessage = (Get-Date -format "yy/MM/dd hh:mm:ss")+", ERROR,   $LogMessage"
        Write-host $LogMessage -foregroundcolor Red
        $LogMessage>> $LogPath
    }
}

# Start Log
"Starting logging for Office365 AD Audit" >> $LogPath
"" >> $LogPath


# Ensure AzureAD and MSOnline Modules are installed
Try{
    Write-Log -LogMessage "Checking Azure modules are installed and localhost is a domain controller..." -LogType Info

    If(!(Get-Module ActiveDirectory -ListAvailable)){
        Write-Log -LogMessage "Localhost is not a Domain Controller or the ActiveDirectory Module is not installed." -LogType Error
        Write-Log -LogMessage ("----------Script Complete------------") -LogType Info
        break
    }
    If(!(Get-Module AzureAD -ListAvailable)){
        Write-Log -LogMessage "Installing AzureAD Module..." -LogType Debug
        Install-module AzureAD -Force
    }
    If(!(Get-Module MSOnline -ListAvailable)){
        Write-log -LogMessage "Installing MSOnline Module..." -LogType Debug
        Install-module MSOnline -Force
    }
    If((Get-Module MSOnline -ListAvailable) -and (Get-Module AzureAD -ListAvailable)){ 
        Write-log -LogMessage "AzureAD and MSOnline modules are installed." -LogType Info
        }
    Else{ Throw "Problem Installing Modules, Check security/firewall settings and connection to internet" }

} 
Catch{
    Write-log -LogMessage "Unable to install azure modules, Restart the script with Administrative privileges and check internet connection." -LogType Error
    Write-Log -LogMessage ("----------Script Complete------------") -LogType Info
    Break
}

$ErrorActionPreference = "SilentlyContinue"

# Login to MSOnline
$Domain = (Get-ADDomain).name
Write-Log -LogMessage "Please enter the administrative credentials for the Office365 Tenant" -LogType ActionRequired
$Credentials =  (Get-Credential ($Domain))
Write-Log -LogMessage "Connecting to MSOnline" -LogType Info
Try{Connect-MsolService -Credential $Credentials -ErrorAction Stop}
Catch{Write-Log -LogMessage "Failed to connect MSOnline" -LogType Error;Write-Log -LogMessage $Error[0] -LogType Error;Write-Log -LogMessage ("----------Script Complete------------") -LogType Info;break}

# Get O365 Licenses and User list
Write-Log -LogMessage "Gathering O365 Users and licenses" -LogType Info
$UserListO365 = Get-MsolUser | Where-Object{$_.FirstName -ne ""} | select Firstname,Lastname,UserprincipalName,Licenses | sort licenses -Descending
$UserListAD = Get-ADUser -Filter{Enabled -eq $true} -properties * | select Userprincipalname

# Get AD Licensing Groups with users in them
Write-Log -LogMessage "Gathering AD Licensing groups" -LogType Info
$DN = (Get-ADDomain).DistinguishedName
$LicensingGroups = Get-ADGroup -SearchBase "OU=Licensing Groups,$DN" -Filter * | Where-Object{((Get-ADGroupMember -identity $_.name).name.length) -gt 0}


# Build Data Table to store information and export
Write-Log -LogMessage "Building information table" -LogType Info
$Table = New-Object system.Data.DataTable “UserListAudit"
$Table.columns.add((New-Object system.Data.DataColumn FirstName,([string])))
$Table.columns.add((New-Object system.Data.DataColumn LastName,([string])))
$Table.columns.add((New-Object system.Data.DataColumn ADUserPrincipalName,([string])))
$Table.columns.add((New-Object system.Data.DataColumn O365UserPrincipalName,([string])))
$Table.columns.add((New-Object system.Data.DataColumn ADLicenses,([string])))
$Table.columns.add((New-Object system.Data.DataColumn O365Licenses,([string])))

# Put information into the Data table
Foreach($User in $UserListO365){
    
    $FirstName = $User.FirstName
    $LastName = $User.LastName

    # Create Row Object
    $Row = $Table.NewRow()

    # Import data from MSOnline pull into the row for the new table
    $Row.FirstName              = $User.Firstname
    $Row.LastName               = $User.LastName      
    $Row.O365UserPrincipalName  = $User.UserPrincipalName
        
    # Attempt to locate the account in AD. If it isn't found by first and lastname match, then it trys to match UPN to display name then it trys just first name, if that fails, it passes and logs
    If((Get-ADUser -Filter {Enabled -eq $True -and GivenName -eq $Firstname -and Surname -eq $LastName}).length -gt 1){Write-Log -LogMessage ("Multiple user accounts detected in AD for "+$User.UserPrincipalName+", There may be potential inaccuracies.") -LogType Error}
    $Row.ADUserPrincipalName  = (Get-ADUser -Filter {Enabled -eq $True -and GivenName -eq $Firstname -and Surname -eq $LastName})[0].UserPrincipalName
        [String]$UPNAttempt = ($User.UserPrincipalName).Split("@")[0]
        If($Row.ADUserPrincipalName.length -lt 2){$Row.ADUserPrincipalName = (Get-ADUser -Filter {Enabled -eq $True -and Name -eq $UPNAttempt}).UserPrincipalName}
        If($Row.ADUserPrincipalName.length -lt 2){$Row.ADUserPrincipalName = (Get-ADUser -Filter {Enabled -eq $True -and GivenName -eq $Firstname}).UserPrincipalName}
        If($Row.ADUserPrincipalName.length -lt 2){Write-Log -LogMessage ($User.UserPrincipalName+" Can't be found in AD.") -LogType Info}

    # Get O365 License Groups
    Write-Log -LogMessage ("Gathering O365 Licenses for "+$User.UserPrincipalName) -LogType info
    # SKip users with 0 licenses in account
    If($User.Licenses.AccountSkuId.Length -lt 1 ){pass}
    # When user has more than 1 license, it is an object and not an array, thus it is added to the table differently
    # Catches users with only 1 license
    ElseIf($User.Licenses.AccountSkuId.GetType().basetype.Name -eq "Object"){
        $LicensePrefix = $User.Licenses.AccountSkuId.split(":")[0]
        $Row.O365Licenses = ($User.Licenses.AccountSkuId).replace("$LicensePrefix"+":","")
    }
    # Catches users with more than 1 license
    ElseIf($User.Licenses.AccountSkuId.GetType().basetype.Name -eq "Array"){
        $Row.O365Licenses = ""
        $O365LicensesArray = @()
        Foreach($Item in $User.Licenses.AccountSkuId){$O365LicensesArray += $Item}
        foreach($item in $O365LicensesArray){
            $LicensePrefix = $Item.split(":")[0]
            $Item = $Item.replace("$LicensePrefix"+":","")
            $Row.O365Licenses += $Item+", "
        }
        $Row.O365Licenses = $Row.O365Licenses.TrimEnd(", ")
    }

    # Get AD License Groups
    $SamAccountName = (Get-ADUser -Properties SamAccountName -Filter {Enabled -eq $True -and GivenName -eq $Firstname -and Surname -eq $LastName}).SamAccountName
    If($SamAccountName.length -lt 1 -and $Row.ADUserPrincipalName.length -lt 2){Write-Log -LogMessage ("Skipping AD License Lookup for " +$User.UserPrincipalName+", Account not in AD or is disabled.") -LogType Info}
    Else{
        Write-Log -LogMessage ("Gathering AD Licensing Groups for "+$User.UserPrincipalName) -LogType Info
        $Row.ADLicenses = ""
        Foreach($Group in $LicensingGroups){
            If((Get-ADGroupMember -Identity $Group.Name | select samaccountname).samaccountname.contains($SamAccountName)){
                $Row.ADLicenses += ($Group.Name + ", ")
            }
        }
        $Row.ADLicenses = $Row.ADLicenses.TrimEnd(", ")
    }
    $Table.Rows.Add($Row)
}

# Output CSV
Write-Log -LogMessage ("Exporting CSV to $Path") -LogType Info
$Output = $Table.rows | sort firstname -Descending | Ft
$Table.rows | export-csv "$Path\$date-O365ADAudit.csv" -NoClobber -NoTypeInformation

Write-Log -LogMessage ("----------Script Complete------------") -LogType Info



