<#	
	.NOTES
	===========================================================================
	 Created on:   	18.12.2019 15:15
     Revised on:    13.12.2020 22:14	 
     Created by:   	Slasky86
     Last revised by: Slasky86
	 Filename:     	New-Admin_Users_and_Safes.ps1
     Version:       0.90
     Last Change:   Changed usage of powershell modules from EPV scripts to psPAS module from psPete
	===========================================================================
	.DESCRIPTION
        This script checks if an admin user is created and creates a personal safe for the unprivileged user
        which is related to the admin user.

        This script must be run on a server where you can reach your AD enviroment and your PVWA.

    Changelog: 

    0.8: Added a check to see if the admin-user already exists
    0.9: Changed from the EPV scripts to the psPAS module from psPete
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    $Global:VaultAdminGroup,
    [Parameter]
    $psPeteModulePath = "",
    $Global:Creds = "",
    $Global:AdminUserPrefix = "adm.",
    $Global:B2BPrefix = "B2B-"
    

)

if ($psPeteModulePath -eq "") {

    $psPeteModulePath = Read-Host "Please write the path of the psPAS module from psPete (psPAS.psd1)"

}

Import-Module $psPeteModulePath

if ($Creds -eq "") {

    $Creds = Get-Credential

}

# Setting a default password for the admin-account
$Global:password = ConvertTo-SecureString -String 'ChangeThis4$ap%%' -AsPlainText -Force

function Create-AdminUser_and_safe {
    

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Creds,
        [Parameter(Mandatory)]
        [string[]]$Users,
        [Parameter(Mandatory)]
        $BaseURI,
        [Parameter(Mandatory)]
        $OrgDomain,
        [Parameter(Mandatory)]
        $PrivUserOU
    )

    New-PASSession -Credential $Creds -BaseURI $BaseURI -type LDAP

    foreach ($User in $Users) {
                
        try {

            $adusers += Get-ADUser $User -ErrorAction Continue

        }

        catch {

            $message = $_.Exception.Message
            Write-Host -ForegroundColor Red "Something went wrong!`n $message"
            Write-Host -ForegroundColor Yellow "Continuing to the next user"

        }
    }
    

    foreach ($user in $adusers) {

        if ($user.SamAccountName -like "$Global:B2BPrefix*") {
            
            $upn = $user.UserPrincipalName

            #Setter adm. prefix på brukernavn før sjekk
            if ($user.SamAccountName -like "$Global:B2BPrefix*") {
                $user = $user.SamAccountName -replace "$Global:B2BPrefix",""
                }

            $userinfo = (Get-Culture).TextInfo
            $usersafe = $userinfo.ToTitleCase($user)
            $adminuser = "$global:AdminUserPrefix"+$user

        }

        else {

            $userinfo = (Get-Culture).TextInfo
            $usersafe = $userinfo.ToTitleCase($usersafe)
            $adminuser = "$global:AdminUserPrefix"+$user.SamAccountName
            $upn = $user.UserPrincipalName

        }

        # Checks if the admin-user exists already, and if not, creates it
        $doesExist = Get-ADUser $adminuser -ErrorAction SilentlyContinue
        Write-Host -ForegroundColor Yellow "Checking if $adminuser exists..."
        if (!$doesExist) {

            $user = get-aduser $user -Properties DisplayName, GivenName, Surname
            [string]$displayname = $user.DisplayName
            [string]$givenname = $user.givenname
            [string]$surname = $user.surname
            New-ADUser -Name $Displayname -SamAccountName $adminuser -DisplayName $DisplayName -GivenName $givenname -Surname $surname -Path $PrivUserOU -UserPrincipalName "$adminuser@$OrgDomain" -Enabled $true -AccountPassword $Global:password -Description "Automatically created by CyberArk script"
            Write-Host -ForegroundColor Green "User $adminuser is created"

            Add-PASSafe -SafeName $usersafe -Description "Created by script" -OLACEnabled $false -ManagingCPM "PasswordManager" -NumberOfVersionsRetention 5
            Add-PASSafeMember -SafeName $SafeName -SearchIn "vault" -MemberName "Administrator" -UseAccounts $true -ListAccounts $true -RetrieveAccounts $true -AddAccounts $true -UpdateAccountContent $true -RenameAccounts $true -UpdateAccountProperties $true -InitiateCPMAccountManagementOperations $true -SpecifyNextAccountContent $true -DeleteAccounts $true -DeleteFolders $true -UnlockAccounts $true -ManageSafe $true -ManageSafeMembers $true -ViewAuditLog $true -ViewSafeMembers $true -CreateFolders $true -MoveAccountsAndFolders $true -AccessWithoutConfirmation $true
            Add-PASSafeMember -SafeName $SafeName -SearchIn $OrgDomain -MemberName $Global:VaultAdminGroup -UseAccounts $true -ListAccounts $true -RetrieveAccounts $true -AddAccounts $true -UpdateAccountContent $true -RenameAccounts $true -UpdateAccountProperties $true -InitiateCPMAccountManagementOperations $true -SpecifyNextAccountContent $true -DeleteAccounts $true -DeleteFolders $true -UnlockAccounts $true -ManageSafe $true -ManageSafeMembers $true -ViewAuditLog $false -ViewSafeMembers $true -CreateFolders $true -MoveAccountsAndFolders $true -AccessWithoutConfirmation $true   
            Add-PASSafeMember -SafeName $safename -SearchIn $OrgDomain -MemberName $upn -UseAccounts $true -ListAccounts $true -RetrieveAccounts $false -AddAccounts $true -UpdateAccountContent $true -RenameAccounts $true -UpdateAccountProperties $true -InitiateCPMAccountManagementOperations $true -SpecifyNextAccountContent $false -DeleteAccounts $false -UnlockAccounts $false -ManageSafe $false -ManageSafeMembers $false -ViewAuditLog $false -ViewSafeMembers $false -CreateFolders $false -MoveAccountsAndFolders $false -AccessWithoutConfirmation $true
            Add-PASAccount -name "Admin account for $user" -userName $adminuser -address  $OrgDomain -SafeName $usersafe -platformID "WinDomain" -automaticManagementEnabled $true
            Remove-PASSafeMember -SafeName $SafeName -MemberName $creds.username

        }

        else {

            Write-host -ForegroundColor Yellow "User $adminuser exists already, will not create..."
    
        }

        $doesExist = $null

    } 
}

