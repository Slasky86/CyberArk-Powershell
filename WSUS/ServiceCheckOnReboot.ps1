## Script to check if the services are running after reboot

$servicename = "PrivateArk Server"

$service = Get-service -Name $servicename

Start-Sleep -Seconds 300

if ($service.Status -ne "Running") {

    try {

        $DRservice = Get-serivce -name "CyberArk Vault Disaster Recovery"

        if ($DRservice.status -ne "Running") {

            Start-Service "PrivateArk Server" -ErrorAction STOP

            }
        }

    catch {

        $errormessage = $_.Exception.message
        $mailbody = "PrivateArk Service has not started on $env:computername. Errormessage: $errormessage! Please advice that the service must be started manually for the CyberArk enviroment to work"
        Send-MailMessage -From '<EnterSenderEmailHere>' -To "<EnterRecipientHere>" -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer '<EnterSMTPServerHere>'

    }
}

if ($service.Status -eq "Running") {

    $mailbody = "PrivateArk Service has started on $env:computername after a reboot of the server"
    Start-Service "Cyber-Ark Event Notification Engine" -ErrorAction STOP
    Send-MailMessage -From '<EnterSenderEmailHere>' -To "<EnterRecipientHere>" -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer '<EnterSMTPServerHere>'

}