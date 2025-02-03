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

$ServerServiceName = "PrivateArk Server"
$DRServiceName = "CyberArk Vault Disaster Recovery"
$ClusterVaultServiceName = "CyberArk Cluster Vault Manager"


$ClusterVaultService = Get-Service -Name $ClusterVaultServiceName -ErrorAction SilentlyContinue
$PrivateArkService = Get-service -Name $ServerServiceName -ErrorAction SilentlyContinue
$DRservice = Get-service -name $DRServiceName -ErrorAction SilentlyContinue
$ClusterVaultService = Get-Service -Name $ClusterVaultServiceName -ErrorAction SilentlyContinue

if ($ClusterVaultService -notin "",$null) {

    try {

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

            if ($ClusterVaultService.Status -eq "Running" -and $PrivateArkService.status -ne "Running" -and $DRservice.status -eq "Running") {

                $mailbody = "Cluster Vault service has started on $env:computername. This is the DR node"
                Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority Normal -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
    
            }
        }
    }

    catch {

        $errormessage = $_.Exception.message
        $mailbody = "PrivateArk Service has not started on $env:computername. Errormessage: $errormessage! Please advice that the service must be started manually for the CyberArk enviroment to work"
        Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
        
    }

}

if ($ClusterVaultService -in "",$null) {

    try {

        if ($DRservice.status -ne "Running" -and $PrivateArkService.status -eq "Running") {

            $mailbody = "PrivateArk Service has started on $env:computername."
            Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
            

        }
    }

    catch {

        $errormessage = $_.Exception.message
        $mailbody = "PrivateArk Service has not started on $env:computername. Errormessage: $errormessage! Check logs and verify service"
        Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
        

    }

    try {

        if ($DRservice.status -ne "Running" -and $PrivateArkService.status -ne "Running") {

            Start-Service "PrivateArk Server" -ErrorAction STOP

        }
    }

    catch {

        $errormessage = $_.Exception.message
        $mailbody = "PrivateArk Service has not started on $env:computername. Errormessage: $errormessage! Check logs and verify service"
        Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
        

    }

    try {

        if ($DRservice.status -eq "Running" -and $PrivateArkService.Status -ne "Running") {

            $mailbody = "Disaster Recovery service has started on $env:computername."
            Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority Normal -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
                
        }
    }

    catch {

        $errormessage = $_.Exception.message
        $mailbody = "One or more of the services on $env:computername is not running as expected. Errormessage: $errormessage! Check logs and verify service"
        Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Service' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
        
    }
}


netsh advfirewall firewall delete rule name=SMTP-Out dir=out | Out-Null