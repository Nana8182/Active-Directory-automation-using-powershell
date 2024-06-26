##########################################
## Automating Active Directory
## Date: June 24, 2024
## By: Naol Teshome
##########################################

# Load in the CSV file for employees
function Get-EmployeeFromCsv {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string]$Delimiter,
        [Parameter(Mandatory)]
        [hashtable]$SyncFieldMap
    )

    try {
        $SyncProperties = $SyncFieldMap.GetEnumerator()
        $Properties = foreach ($Property in $SyncProperties) {
            @{ Name = $Property.Value; Expression = [scriptblock]::Create("`$_.$($Property.Key)") }
        }

        Import-Csv -Path $FilePath -Delimiter $Delimiter | Select-Object -Property $Properties
    } catch {
        Write-Error $_.Exception.Message
    }
}

# Load in the employees already in AD
function Get-EmployeesFromAD {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [hashtable]$SyncFieldMap,
        [Parameter(Mandatory)]
        [string]$Domain,
        [Parameter(Mandatory)]
        [string]$UniqueID
    )

    try {
        Get-ADUser -Filter {$UniqueID -like "*"} -Server $Domain -Properties @($SyncFieldMap.Values)
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Compare those
function Compare-Users {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [hashtable]$SyncFieldMap,
        [Parameter(Mandatory)]
        [string]$UniqueID,
        [Parameter(Mandatory)]
        [string]$CSVFilePath,
        [Parameter()]
        [string]$Delimiter = ",",
        [Parameter(Mandatory)]
        [string]$Domain
    )

    try {
        $CSVUsers = Get-EmployeeFromCsv -FilePath $CSVFilePath -Delimiter $Delimiter -SyncFieldMap $SyncFieldMap
        $ADUsers = Get-EmployeesFromAD -SyncFieldMap $SyncFieldMap -UniqueID $UniqueID -Domain $Domain

        Compare-Object -ReferenceObject $ADUsers -DifferenceObject $CSVUsers -Property $UniqueID -IncludeEqual
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Get the new users, synced users, and removed users
function Get-UserSyncData {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [hashtable]$SyncFieldMap,
        [Parameter(Mandatory)]
        [string]$UniqueID,
        [Parameter(Mandatory)]
        [string]$CSVFilePath,
        [Parameter()]
        [string]$Delimiter = ",",
        [Parameter(Mandatory)]
        [string]$Domain,
        [Parameter(Mandatory)]
        [string]$OUProperty
    )

    try {
        $CompareData = Compare-Users -SyncFieldMap $SyncFieldMap -UniqueID $UniqueID -CSVFilePath $CSVFilePath -Delimiter $Delimiter -Domain $Domain
        $NewUsersID = $CompareData | Where-Object { $_.SideIndicator -eq "=>" }
        $SyncedUsersID = $CompareData | Where-Object { $_.SideIndicator -eq "==" }
        $RemovedUsersID = $CompareData | Where-Object { $_.SideIndicator -eq "<=" }

        $NewUsers = Get-EmployeeFromCsv -FilePath $CSVFilePath -Delimiter $Delimiter -SyncFieldMap $SyncFieldMap | Where-Object { $UniqueID -in $NewUsersID.$UniqueID }
        $SyncedUsers = Get-EmployeeFromCsv -FilePath $CSVFilePath -Delimiter $Delimiter -SyncFieldMap $SyncFieldMap | Where-Object { $UniqueID -in $SyncedUsersID.$UniqueID }
        $RemovedUsers = Get-EmployeesFromAD -SyncFieldMap $SyncFieldMap -Domain $Domain -UniqueID $UniqueID | Where-Object { $UniqueID -in $RemovedUsersID.$UniqueID }

        @{
            New = $NewUsers
            Synced = $SyncedUsers
            Removed = $RemovedUsers
            Domain = $Domain
            UniqueID = $UniqueID
            OUProperty = $OUProperty
        }
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Generate a new username
function New-UserName {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$GivenName,
        [Parameter(Mandatory)]
        [string]$Surname,
        [Parameter(Mandatory)]
        [string]$Domain
    )

    try {
        [RegEx]$Pattern = "\s|-|'"
        $index = 1

        do {
            $Username = "$Surname$($GivenName.Substring(0, $index))" -replace $Pattern, ""
            $index++
        } while ((Get-ADUser -Filter "SamAccountName -like '$Username'" -Server $Domain) -and ($Username -notlike "$Surname$GivenName"))

        if (Get-ADUser -Filter "SamAccountName -like '$Username'" -Server $Domain) {
            throw "No usernames available for this user!"
        } else {
            $Username
        }
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Validate organizational units
function Validate-OU {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [hashtable]$SyncFieldMap,
        [Parameter(Mandatory)]
        [string]$CSVFilePath,
        [Parameter()]
        [string]$Delimiter = ",",
        [Parameter(Mandatory)]
        [string]$Domain,
        [Parameter()]
        [string]$OUProperty
    )

    try {
        $OUNames = Get-EmployeeFromCsv -FilePath $CSVFilePath -Delimiter $Delimiter -SyncFieldMap $SyncFieldMap | Select-Object -Unique -Property $OUProperty

        foreach ($OUName in $OUNames) {
            $OUName = $OUName.$OUProperty
            if (-not (Get-ADOrganizationalUnit -Filter "name -eq '$OUName'" -Server $Domain)) {
                New-ADOrganizationalUnit -Name $OUName -Server $Domain -ProtectedFromAccidentalDeletion $false
            }
        }
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Create new users
function Create-NewUsers {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [hashtable]$UserSyncData
    )

    try {
        $NewUsers = $UserSyncData.New

        foreach ($NewUser in $NewUsers) {
            Write-Verbose "Creating user: $($NewUser.GivenName) $($NewUser.Surname)"
            $Username = New-UserName -GivenName $NewUser.GivenName -Surname $NewUser.Surname -Domain $UserSyncData.Domain
            Write-Verbose "Creating user: $($NewUser.GivenName) $($NewUser.Surname) with username: $Username"

            if (-not ($OU = Get-ADOrganizationalUnit -Filter "name -eq '$($NewUser.$($UserSyncData.OUProperty))'" -Server $UserSyncData.Domain)) {
                throw "The organizational unit {$($NewUser.$($UserSyncData.OUProperty))} does not exist."
            }

            Write-Verbose "Creating user: $($NewUser.GivenName) $($NewUser.Surname) with username: $Username, $OU"

            Add-Type -AssemblyName 'System.Web'
            $Password = [System.Web.Security.Membership]::GeneratePassword((Get-Random -Minimum 12 -Maximum 15), 3)
            $SecuredPassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

            $NewADUserParams = @{
                EmployeeID = $NewUser.EmployeeID
                GivenName = $NewUser.GivenName
                Surname = $NewUser.Surname
                Name = $Username
                SamAccountName = $Username
                UserPrincipalName = "$Username@$($UserSyncData.Domain)"
                AccountPassword = $SecuredPassword
                ChangePasswordAtLogon = $true
                Enabled = $true
                Title = $NewUser.Title
                Department = $NewUser.Department
                Office = $NewUser.Office
                Path = $OU.DistinguishedName
                Confirm = $false
                Server = $UserSyncData.Domain
            }

            New-ADUser @NewADUserParams
            Write-Verbose "Created user: $($NewUser.GivenName) $($NewUser.Surname) EmpID: $($NewUser.EmployeeID) Username: $Username Password: $Password"
        }
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Check username and generate if needed
function Check-UserName {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$GivenName,
        [Parameter(Mandatory)]
        [string]$Surname,
        [Parameter(Mandatory)]
        [string]$CurrentUserName,
        [Parameter(Mandatory)]
        [string]$Domain
    )

    try {
        if ((Get-ADUser -Filter "SamAccountName -like '$CurrentUserName'" -Server $Domain)) {
            $Username = New-UserName -GivenName $GivenName -Surname $Surname -Domain $Domain
        } else {
            $Username = $CurrentUserName
        }

        $Username
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Sync the data for the existing users
function Sync-ExistingUsers {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [hashtable]$UserSyncData,
        [Parameter(Mandatory)]
        [hashtable]$SyncFieldMap
    )

    try {
        $SyncedUsers = $UserSyncData.Synced

        foreach ($SyncedUser in $SyncedUsers) {
            $ADUser = Get-ADUser -Filter "$($UserSyncData.UniqueID) -eq '$($SyncedUser.$($UserSyncData.UniqueID))'" -Server $UserSyncData.Domain -Properties @($SyncFieldMap.Values)
            $Username = Check-UserName -GivenName $SyncedUser.GivenName -Surname $SyncedUser.Surname -CurrentUserName $ADUser.SamAccountName -Domain $UserSyncData.Domain

            $Params = @{
                Identity = $ADUser
                SamAccountName = $Username
                UserPrincipalName = "$Username@$($UserSyncData.Domain)"
                Title = $SyncedUser.Title
                Department = $SyncedUser.Department
                Office = $SyncedUser.Office
                Confirm = $false
                Server = $UserSyncData.Domain
            }

            if ($OU = Get-ADOrganizationalUnit -Filter "name -eq '$($SyncedUser.$($UserSyncData.OUProperty))'" -Server $UserSyncData.Domain) {
                $Params["MoveTo"] = $OU.DistinguishedName
            }

            Set-ADUser @Params
            Write-Verbose "Synced user: $($SyncedUser.GivenName) $($SyncedUser.Surname) with username: $Username"
        }
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Disable the users not in the current CSV file
function Remove-Users {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [hashtable]$UserSyncData,
        [Parameter()]
        [int]$KeepDisabledForDays = 7
    )

    try {
        $RemovedUsers = $UserSyncData.Removed

        foreach ($RemovedUser in $RemovedUsers) {
            $Params = @{
                Identity = $RemovedUser
                AccountExpirationDate = (Get-Date).AddDays($KeepDisabledForDays)
                Confirm = $false
                Server = $UserSyncData.Domain
            }

            Set-ADUser @Params
            Disable-ADAccount -Identity $RemovedUser
            Write-Verbose "Disabled user: $($RemovedUser.GivenName) $($RemovedUser.Surname)"
        }
    } catch {
        Write-Error -Message $_.Exception.Message
    }
}

# Validate the organizational units in AD
Validate-OU -SyncFieldMap $SyncFieldMap -CSVFilePath $CSVFilePath -Delimiter $Delimiter -Domain $Domain -OUProperty $OUProperty

# Get the new, synced, and removed user data
$UserSyncData = Get-UserSyncData -SyncFieldMap $SyncFieldMap -UniqueID $UniqueID -CSVFilePath $CSVFilePath -Delimiter $Delimiter -Domain $Domain -OUProperty $OUProperty

# Create the new users
Create-NewUsers -UserSyncData $UserSyncData

# Sync the existing users
Sync-ExistingUsers -UserSyncData $UserSyncData -SyncFieldMap $SyncFieldMap

# Remove the users that are not in the CSV file
Remove-Users -UserSyncData $UserSyncData -KeepDisabledForDays $KeepDisabledForDays

