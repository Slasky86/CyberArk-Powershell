# BaseURL of PVWA without the /PasswordVault
$PVWAURL = "<PVWAURL>"

# Username and password of account that has permissions to view the accounts
$Username = ""
$Password = ""
$usernames = @("")
$csvpath = "C:\Temp"

$Password = ConvertTo-SecureString $Password -AsPlainText -Force


# Creating the credentials object for authentication
$Credential = New-Object System.Management.Automation.PSCredential($Username,$Password)

$body = @{

    "username"= $Credential.UserName;
    "password"= ([System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password)));
    "concurrentSession"= "true"

}

# Retrieving an authentication token
$authentication = Invoke-RestMethod "$PVWAURL/PasswordVault/API/Auth/CyberArk/Logon" -Method Post -Body ($body | ConvertTo-Json) -ContentType application/json -UseBasicParsing

$authbody = @{

    "Authorization"="$authentication"

}

$safes = Invoke-RestMethod -Uri "$PVWAURL/PasswordVault/api/safes" -Headers $authbody -ContentType application/json -UseBasicParsing

foreach ($safe in $safes.value) {

    foreach ($username in $usernames) {

        $searchbody = @{

            username = $username

        }

        $user = Invoke-RestMethod -Uri "$PVWAURL/PasswordVault/API/Users/" -Headers $authbody -Body $searchbody -ContentType Application/json -UseBasicParsing

        $userMembership = Invoke-RestMethod -Uri "$PVWAURL/PasswordVault/API/Users/$($user.Users.id)" -Headers $authbody -ContentType Application/json -UseBasicParsing | select -ExpandProperty groupsMembership

        $safemembers = Invoke-RestMethod -Uri "$PVWAURL/PasswordVault/api/Safes/$($Safe.safeUrlId)/Members/" -Headers $authbody -ContentType Application/json -UseBasicParsing | select -ExpandProperty Value

        $safepermissions = $safemembers | where {$_.memberName -in $($userMembership.groupname) -or $_.memberName -eq $username} | select memberName, memberType, safeName, Permissions

        $safepermissions | Export-Csv -Path $csvpath\$username.csv -Append -NoTypeInformation -Encoding UTF8 -Delimiter ";"

        $safepermissions = $null

    }
}