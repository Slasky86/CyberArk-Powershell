$PVWAURL = "<PVWAURL>"

$Username = "Administrator"
$Password = Read-Host "Provide a password:"

$body = @{
	"username"= "$($Username)";
	"password"= "$($Password)";
	#"newPassword": "{{pasNewPassword}}";
	"concurrentSession"= "true"
}


$authentication = Invoke-RestMethod "$PVWAURL/PasswordVault/API/Auth/CyberArk/Logon" -Method Post -Body ($body | ConvertTo-Json) -ContentType application/json

$authbody = @{

    "Authorization"="$authentication"

}

    
$AllSafes = Invoke-RestMethod "$PVWAURL/PasswordVault/api/Safes?search=BZ&limit=200&extendedDetails=true" -Method Get -ContentType application/json -Headers $authbody
#$AllSafesNames = $AllSafes.value.safename

$BZSafes = $AllSafes | where {$_.SafeName -like "BZ*"}

$i = 0

foreach ($safe in $AllSafes.value) {
    
    Write-Host "Touching $($Safe.safename) which is $($safe.description)"
    $response = Invoke-RestMethod "$PVWAURL/PasswordVault/api/Safes/$($safe.safename)/Members" -Method Get -ContentType application/json  -Headers $authbody 
    $response.value
    $i++
    $i
}

$safename = Read-Host "Write the safename of the safe you want to search for"

$searchbody = @{

    "search"="$($safename)";
    "limit" = "1000"

}

$safes = Invoke-RestMethod "$PVWAURL/PasswordVault/api/Safes/" -Body $searchbody -ContentType application/json -Headers $authbody

foreach ($safe in $safes.value) {

    $accounts = Invoke-RestMethod "$PVWAURL/PasswordVault/api/accounts?filter=safename eq $($safe.safeName)" -ContentType application/json -Headers $authbody
    $accounts.value | select userName, safeName, address, platformID, name | export-csv <export path> -NoTypeInformation -Encoding UTF8 -Append


}