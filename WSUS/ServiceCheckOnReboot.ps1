## Script to check if the services are running after reboot
[CmdletBinding()]
param (
	[Parameter(Mandatory)]
	[string[]]$Recipient,
	[Parameter(Mandatory)]
	$SendFrom,
	[Parameter(Mandatory)]
	$SMTPServer
)


$servicename = "PrivateArk Server"

$service = Get-service -Name $servicename

Start-Sleep -Seconds 300

if ($service.Status -ne "Running") {

    try {

        $DRservice = Get-serivce -name "CyberArk Vault Disaster Recovery"

        if ($DRservice.status -ne "Running") {

            Start-Service "PrivateArk Server" -ErrorAction STOP

            }

        if ($DRservice.status -eq "Running") {

            $mailbody = "Disaster Recovery service has started on $env:computername."
            Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority NOrmal -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer

        }
    }

    catch {

        $errormessage = $_.Exception.message
        $mailbody = "PrivateArk Service has not started on $env:computername. Errormessage: $errormessage! Please advice that the service must be started manually for the CyberArk enviroment to work"
        Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer

    }
}

if ($service.Status -eq "Running") {

    $mailbody = "PrivateArk Service has started on $env:computername after a reboot of the server"
    Start-Service "Cyber-Ark Event Notification Engine" -ErrorAction STOP
    Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer

}