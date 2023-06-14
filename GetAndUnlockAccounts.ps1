
# Enter your PVWA base URL here, without /PasswordVault
$PVWAURL = "<PVWAURL>"

# Credentials for user with the correct permissions to search in safes and unlock accounts
$Global:Creds = Get-Credential


# Function for retrieving accounts and potentially unlock them
function Unlock-PASAccount {

    [CmdletBinding(DefaultParameterSetName="Safename")]
    param (
	[Parameter(Mandatory, ParameterSetName="Safename")]
	[string]$Safename,
    [Parameter(Mandatory, ParameterSetName="All")]
	[switch]$All,
    [Parameter()]
    [switch]$Unlock
    )

    # Start a new PAS Session
    New-PASSession -Credential $creds -BaseURI $PVWAURL -concurrentSession $true

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
            "username"= $Creds.UserName;
            "password"= ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Creds.Password)));
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
Unlock-PASAccount -All -Unlock