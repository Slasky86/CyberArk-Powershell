## Script to update safe retention setting for a number of days

[CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        $PVWAURL,
        [Parameter(Mandatory)]
        $Safelist,
        [Parameter()]
        [string]$DaysOfRetention = "7"
    )

$Credential = Get-Credential

try {

    $Safes = Import-Csv $Safelist -Header "Safename" -Delimiter ";" -ErrorAction Stop

}

catch {

    Write-Host -ForegroundColor Red "Something went wrong while importing the CSV file:"
    $exception.message

}

$body = @{

    "username"= $Credential.UserName;
    "password"= ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)));
    "concurrentSession"= "true"

}

$authentication = Invoke-RestMethod "$PVWAURL/PasswordVault/API/Auth/CyberArk/Logon" -Method Post -Body ($body | ConvertTo-Json) -ContentType application/json -UseBasicParsing

$authbody = @{

    "Authorization"="$authentication"

}


foreach ($Safe in $Safes.Safename) {

    $searchbody = @{

        "search"="$safe"

    }

    $Safedetails = Invoke-RestMethod "$PVWAURL/PasswordVault/API/Safes/" -Method Get -Headers $authbody -Body $searchbody | select -ExpandProperty Value

    $safedetails.numberOfDaysRetention = $DaysOfRetention

    Invoke-RestMethod "$PVWAURL/PasswordVault/API/Safes/$($Safedetails.safeUrlId)" -Method Put -Headers $authbody -Body ($Safedetails | ConvertTo-Json) -UseBasicParsing -ContentType application/json

}
