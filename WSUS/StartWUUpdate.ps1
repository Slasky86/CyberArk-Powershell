## Script that starts and runs Windows Update automatically

[CmdletBinding()]
param (
	[Parameter(Mandatory)]
	[string[]]$Recipient,
	[Parameter(Mandatory)]
	$SendFrom,
	[Parameter(Mandatory)]
	$SMTPServer
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
if ($NeedsReboot -eq $true) {

Restart-Computer

}