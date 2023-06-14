function Unlock-LockedPASAccount {

    [CmdletBinding(DefaultParameterSetName="Safename")]
    param (
	[Parameter(Mandatory, ParameterSetName="Safename")]
	[string]$Safename,
    [Parameter(Mandatory, ParameterSetName="All")]
	[switch]$All,
    [Parameter(Mandatory)]
    [string]$PVWAURL,
    [Parameter(Mandatory)]
    [PSCredential]$Credential,
    [Parameter()]
    [switch]$Unlock
    )

<#	

.NOTES
    ===========================================================================
    Created on:   	13.06.23
    Created by:   	Slasky86
    Organization: 	
    Filename:     	GetAndUnlockAccounts.ps1
    Version:        0.8
    ===========================================================================
    
.SYNOPSIS
    This function searches for locked accounts and unlocks them.

.DESCRIPTION
    This function is made to retrieve accounts, based on safe or all
    available accounts. It utilizes the psPAS powershell module for most
    of the operations, except the unlocking operation, which utilizes an
    API thats not readily available in the main branch of psPAS at this moment.	
            

.CHANGELOG
    Version 0.1
    * Initial creation
    * Function to find all locked accounts

    Version 0.5
    * Added option to search by safe
    * Added unlock option

    Version 0.8
    * Finetuning the function
    * Adding proper parameters
    * Making this fancy thing

.EXAMPLE
PS> Unlock-Account -Credential $Credential -PVWAURL $PVWAURL -All
or
PS> Unlock-Account -Credential (Get-Credential) -PVWAURL "https://pvwa.domain.com" -All

Retrieves all accounts that are locked

.EXAMPLE
PS> Unlock-Account -Credential $Credential -PVWAURL $PVWAURL -All -Unlock
or
PS> Unlock-Account -Credential (Get-Credential) -PVWAURL "https://pvwa.domain.com" -All -Unlock

Retrieves all accounts that are locked, and unlocks them

.EXAMPLE
PS> Unlock-Account -Credential $Credential -PVWAURL $PVWAURL -SafeName $Safename
or
PS> Unlock-Account -Credential (Get-Credential) -PVWAURL "https://pvwa.domain.com" -Safename "DemoSafe"

Retrieves all accounts in the defined safe

.EXAMPLE
PS> Unlock-Account -Credential $Credential -PVWAURL $PVWAURL -SafeName $Safename -Unlock
or
PS> Unlock-Account -Credential (Get-Credential) -PVWAURL "https://pvwa.domain.com" -Safename "DemoSafe" -Unlock

Retrieves all accounts in the defined safe and unlocks them

#>

    # Check if psPAS is installed, and if not, install it

    try {

        $psPAS = Get-Module -ListAvailable -Name "psPAS"

        if ($psPAS -in "",$null) {

            Write-Host -ForegroundColor Green "Trying to install the psPAS powershell module..."
            Install-Module -Name "psPAS" -Scope CurrentUser -AllowClobber -Force
            Write-Host -ForegroundColor Green "Installation successful!"

        }
    }

    catch {

        Write-Error $_.Exception.Message
        Write-Error "Module installation failed. Script will now exit"
        Exit

    }

    # Start a new PAS Session
    New-PASSession -Credential $Credential -BaseURI $PVWAURL -concurrentSession $true

    # Arraylist for Account details of retrieved PAS Accounts
    $AccountsDetails = New-object System.Collections.ArrayList

    # If the -All switch is set, retrieve all accounts
    if ($all) {
        
        Write-Warning "This can take quite some time in a large environment"

        # Retrieving all accounts
        $Accounts = Get-PASAccount
        
        # Iterating through retrieved accounts to get account details    
        foreach ($account in $accounts) {

            $Accountsdetails.add((Get-PASAccountDetail -id $account.id)) | Out-Null
        }        
    }


    else {
        
        # If -All isnt defined, search by safename
        $accounts = Get-PASAccount -safeName $Safename

        # Iterating through retrieved accounts to get account details 
        foreach ($account in $accounts) {

            $Accountsdetails.add((Get-PASAccountDetail -id $account.id)) | Out-Null
        
        }
    }

    # Arraylist for Locked accounts info
    $LockedAccountsInfo = @()

    # Iterate through all retrieved accounts with account details
    foreach ($AccountDetail in $AccountsDetails.Details) {
        
        # Check if account is locked or not
        if ($AccountDetail.LockedBy -notin "",$null) {
            
            # If locked, add information to the array for ease of reading. Attributes can be added if desired
            $LockedAccountsInfo += @{
    
                Username = $AccountDetail.RequiredProperties.Username
                AccountID = $Accounts | where {$_.name -eq $AccountDetail.name} | select -ExpandProperty Id
                Safe = $AccountDetail.safeName
            }
        }    
    }

    # If the unlock switch is set
    if ($Unlock) {
        
        # Authentication body for the REST Call
        $body = @{
            "username"= $Credential.UserName;
            "password"= ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)));
            "concurrentSession"= "true"
        }

        # Retrieving an authentication token 
        $authentication = Invoke-RestMethod "$PVWAURL/PasswordVault/API/Auth/CyberArk/Logon" -Method Post -Body ($body | ConvertTo-Json) -ContentType application/json

        $authbody = @{

            "Authorization"="$authentication"

        }

        # Iterate through all locked accounts
        foreach ($LockedAccount in $LockedAccountsInfo) {

            # Try to unlock accounts one by one, throwing an error message if it fails
            try {
            
                Invoke-RestMethod "$PVWAURL/PasswordVault/api/accounts/$($LockedAccount.AccountID)/unlock" -Method Post -Headers $authbody -ContentType application/json
                Write-Host -ForegroundColor Green "Unlocked $($LockedAccount.userName)"

            }

            catch {

                Write-host -ForegroundColor Red "Error: $_.Exception.Message. Failed to unlock account $($LockedAccount.username)"
                Continue

            }
        }
    }

    # Return information about locked accounts if they arent unlocked directly
    else {
    
        Return $LockedAccountsInfo

    }     
}

# Command to run the function
# Unlock-LockedPASAccount -All -Unlock

