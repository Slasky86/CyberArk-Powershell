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

netsh advfirewall firewall add rule name=SMTP-Out dir=out action=allow protocol=TCP remoteport=25 remoteip=$SMTPServer

Start-Sleep -Seconds 300

$Servicename = "PrivateArk Server"

$PrivateArkService = Get-service -Name $servicename

try {

    $DRservice = Get-service -name "CyberArk Vault Disaster Recovery"
    $ClusterVaultService = Get-Service -Name "CyberArk Cluster Vault Manager"

    if ($ClusterVaultService.Status -eq "Running") {

        if ($PrivateArkService.Status -eq "Running") {

            $mailbody = "PrivateArk Service has started on $env:computername after a reboot of the server. This is the primary node"
            Start-Service "Cyber-Ark Event Notification Engine" -ErrorAction STOP
            Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
            break
        }
        
        if ($PrivateArkService.status -ne "Running" -and $DRservice.Status -ne "Running" -and $ClusterVaultService.Status -eq "Running") {

            $mailbody = "Cluster Vault Manager has started on $env:computername after a reboot of the server. This is the standby node in the cluster."
            Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority Normal -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
           
        }
    }

    if ($DRservice.status -ne "Running" -and $ClusterVaultService -in "",$null) {

        Start-Service "PrivateArk Server" -ErrorAction STOP

    }

    if ($DRservice.status -eq "Running" -and $ClusterVaultService.Status -ne "Running") {

        $mailbody = "Disaster Recovery service has started on $env:computername."
        Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority Normal -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
            
    }

    if ($ClusterVaultService -notin "",$null -and $ClusterVaultService.Status -eq "Running" -and $PrivateArkService.status -ne "Running" -and $DRservice.status -eq "Running") {

        $mailbody = "Cluster Vault service has started on $env:computername. This is the DR node"
        Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority Normal -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer

    }
}

catch {

    $errormessage = $_.Exception.message
    $mailbody = "PrivateArk Service has not started on $env:computername. Errormessage: $errormessage! Please advice that the service must be started manually for the CyberArk enviroment to work"
    Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
       
}


netsh advfirewall firewall delete rule name=SMTP-Out dir=out | Out-Null