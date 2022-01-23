### This function creates a new safe based on input by the user
### This function can create an AD group to manage the specific safe
### Certain parameters might have to be tweaked, like safe permissions, managing CPM etc


function New-Safe {
    
    # This function requires the powershell-module from psPete

    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $Creds,
        [Parameter(Mandatory)]
        $BaseURI,
        [Parameter(Mandatory)]
        $OrgPrefix,
        [Parameter(Mandatory)]
        $GroupPrefix,
        [Parameter(Mandatory)]
        $OrgDomain
    )
    
    $safenames = @()

    While ($true) {

        $safenames += Read-Host "Write the safename, without the OrgPrefix"

        $done = Read-Host "Do you want to create more safes? (Y)es or (N)o"

        if ($done.tolower() -eq "y") {

            continue

        }

        if ($done.tolower() -eq "n") {

            break

        }

        else {

            Write-Host -ForegroundColor Yellow "Please choose Y for Yes and N for No"

        }
    }

    # Connect to PVWA using LDAP credentials
    New-PASSession -Credential $Creds -BaseURI $BaseURI -type LDAP

    foreach ($safename in $safenames) {

        $safe = $OrgPrefix + $safename
        $primaryMember = $GroupPrefix + $safename
        

        Add-PASSafe -SafeName $safe -Description "Created by script" -OLACEnabled $false -ManagingCPM "PasswordManager" -NumberOfVersionsRetention 5

        # Uncomment this and add your own logic to create an AD group to gain access to this safe
        #New-ADGroup -Name $primaryMember -SamAccountName $primaryMember -DisplayName $primaryMember -Path "<BaseDN for your CyberArk groups in AD>" -Description "Accessgroup for safe $safe"

        Add-PASSafeMember -SafeName $Safe -SearchIn "vault" -MemberName "Administrator" -UseAccounts $true -ListAccounts $true -RetrieveAccounts $true -AddAccounts $true -UpdateAccountContent $true -RenameAccounts $true -UpdateAccountProperties $true -InitiateCPMAccountManagementOperations $true -SpecifyNextAccountContent $true -DeleteAccounts $true -DeleteFolders $true -UnlockAccounts $true -ManageSafe $true -ManageSafeMembers $true -ViewAuditLog $true -ViewSafeMembers $true -CreateFolders $true -MoveAccountsAndFolders $true -AccessWithoutConfirmation $true
        Add-PASSafeMember -SafeName $Safe -SearchIn $OrgDomain -MemberName "<your vault admin group here>" -UseAccounts $true -ListAccounts $true -RetrieveAccounts $true -AddAccounts $true -UpdateAccountContent $true -RenameAccounts $true -UpdateAccountProperties $true -InitiateCPMAccountManagementOperations $true -SpecifyNextAccountContent $true -DeleteAccounts $true -DeleteFolders $true -UnlockAccounts $true -ManageSafe $true -ManageSafeMembers $true -ViewAuditLog $false -ViewSafeMembers $true -CreateFolders $true -MoveAccountsAndFolders $true -AccessWithoutConfirmation $true   
        Add-PASSafeMember -SafeName $safe -SearchIn $OrgDomain -MemberName $primaryMember -UseAccounts $true -ListAccounts $true -RetrieveAccounts $false -AddAccounts $true -UpdateAccountContent $true -RenameAccounts $true -UpdateAccountProperties $true -InitiateCPMAccountManagementOperations $true -SpecifyNextAccountContent $false -DeleteAccounts $false -UnlockAccounts $false -ManageSafe $false -ManageSafeMembers $false -ViewAuditLog $false -ViewSafeMembers $false -CreateFolders $false -MoveAccountsAndFolders $false -AccessWithoutConfirmation $true
        Remove-PASSafeMember -SafeName $Safe -MemberName $creds.username
    
    }
}