## Script that starts and runs Windows Update automatically

[CmdletBinding()]
param (
	[Parameter()]
	[string[]]$Recipient = "",
	[Parameter()]
	$SendFrom = "",
	[Parameter()]
	$SMTPServer = "",
	[Parameter()]
	$automaticReboot = $false
)


# Moves current working directory to where the scriptfiles are
cd <#path to script files#>


# Opens the necessary services for Windows update to run
.\OpeningServices.ps1

# Starts the download from the WSUS server
.\DownloadUpdatesFromWSUS.ps1 -Recipient $Recipient -SendFrom $SendFrom -SMTPServer $SMTPServer

# Starts the installation of the updates and retrieves the RebootRequired status
$NeedsReboot = .\InstallUpdates.ps1 -Recipient $Recipient -SendFrom $SendFrom -SMTPServer $SMTPServer

# Closes the services necessary to run Windows Update
.\ClosingServices.ps1

# Restarts the server if the $NeedsReboot variable is $true 
if ($NeedsReboot -eq $true -and $automaticReboot -eq $true) {

    Restart-Computer

}

if ($NeedsReboot -eq $true -and $automaticReboot -eq $false) {

	Write-Warning -Message "A reboot is required for the updates to finish"

}

else {

	Write-Warning -Message "Updates finished installing, no reboot is required"

}