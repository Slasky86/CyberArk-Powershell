$PVWAURL = "<PVWAURL>"

$Creds = Get-Credential

$body = @{
    "username"= $Creds.UserName;
    "password"= ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.Password)));
    "concurrentSession"= "true"
}

function GetLockedAccounts {

    [CmdletBinding()]
    param (
	[Parameter()]
	[string]$Safename="",
	[Parameter()]
	[switch]$All,
    [Parameter()]
    [switch]$Unlock
    )

    New-PASSession -Credential $creds -BaseURI $PVWAURL -concurrentSession $true

    $AccountsDetails = New-object System.Collections.ArrayList

    if ($all) {
        
        Write-Warning "This can take quite some time in a large environment"

        $Accounts = Get-PASAccount
            
        foreach ($account in $accounts) {

            $Accountsdetails.add((Get-PASAccountDetail -id $account.id)) | Out-Null
        }        
    }


    else {

        $accounts = Get-PASAccount -safeName $Safename

        foreach ($account in $accounts) {

            $Accountsdetails.add((Get-PASAccountDetail -id $account.id)) | Out-Null
        
        }
    }

    $LockedAccountsInfo = @()

    foreach ($AccountDetail in $AccountsDetails.Details) {
        
        if ($AccountDetail.LockedBy -notin "",$null) {

            $LockedAccountsInfo += @{
    
                Username = $AccountDetail.RequiredProperties.Username
                AccountID = $Accounts | where {$_.Username -eq $AccountDetail.RequiredProperties.Username} | select -ExpandProperty Id
                Safe = $AccountDetail.safeName
            }
        }    
    }

    if ($Unlock) {

        $authentication = Invoke-RestMethod "$PVWAURL/PasswordVault/API/Auth/CyberArk/Logon" -Method Post -Body ($body | ConvertTo-Json) -ContentType application/json

        $authbody = @{

            "Authorization"="$authentication"

        }

        foreach ($LockedAccount in $LockedAccountsInfo) {

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

    else {
    
        Return $LockedAccountsInfo

    }     
}


GetLockedAccounts -All -Unlock