# WSUS updates on CyberArk on-prem vaults

These scripts are the ones supplied by CyberArk, with slight modifications. Due to these modifications the signature will be invalid and it will fail a signature check. 
The scripts that are not of CyberArk origin are simple scripts for starting the scripts in order and to check if the services are running properly after a reboot.
The scripts sends email notifications as well as write to the Windows Event Viewer.

## Prerequisites

For the Windows Event Viewer logging to work, you need to create a new log source. This is done by typing this in an administrative powershell window on each vault the scripts run on:

    New-EventLog -LogName Application -Source "VaultWUUpdate"
  
You will also have to specify the path of the scriptfiles, as the StartWUUpdate.ps1 script navigates to that folder. This can be circumvented by specifying "Start in" section of the task scheduler, or starting the script from the folder its resides in.

## What does the scripts do?

The first script that should be run is the "ConfigureWSUS.ps1" script. This defines in registry which WSUS server the vault is to connect to. A couple of things to consider here:

1. If the WSUS requires https, the vault needs the Root CA certificate of the WSUS server to be able to trust it. 
2. If the Vault uses a FQDN to connect to the WSUS, add that entry to the hosts file (as per CyberArk recommendations). 
3. If you type the WSUS servername wrong, simply run the script again, with the correct name.

The syntax is:

    .\ConfigureWSUS.ps1 "http://CompanyWSUS:8530"

The scripts in general does the following when run in the correct order:

1. Starts the services needed to be able to run Windows Updates
2. Download the updates from the defined WSUS server
3. Installs the downloaded updates and logs how many are successful and how many failed
4. Sends an email of the statistics and writes to the Event Viewer
5. Checks if the server has a pending reboot
6. Stops the services and removes any temporary firewall rules put in to allow access to SMTP server or WSUS server
7. Reboots the server if automatic reboot is selected. Otherwise it will just print to console and event viewer that the server needs a reboot
8. Checks the CyberArk services for "Running" status after reboot, and sends a mail if something is wrong.


## How to set up these scripts

You can do this in one of two ways:

1. Running the scripts manually at will and when needed
2. Set up a scheduled task for running the scripts at given intervals.

If you run the scripts with scheduled tasks, I recommend setting up two tasks. One that runs the StartWUUpdate.ps1 script and one that runs the ServiceCheckOnReboot.ps1 script.

The task running the update scripts should be set to match the organizations update strategy. My recommendation in DR or Cluster Vault setups is to do a asynchrounus update of the vaults. This is to prevent updating two nodes at the same time, in case the Windows updates break anything. If they do, you will have a node to fail over to. If you got a Cluster Vault environment with a DR site, I recommend updating one node in each site, and do the other node the following week.

The ServiceCheckOnReboot.ps1 script should be set up to run at each startup of the machine, and whether or not a user has logged on.


## Optional parameters

The scripts have a few required parameters, and if they are supplied with the StartWUUpdate.ps1 script it will propagate to the InstallUpdates.ps1 script. The parameters are as follows:

1. Recipient - A list of recipient emails, separated by comma. Ex: "recipient1@company.com","recipient2@company.com"
2. SendFrom - The address that the scripts present themselves as. Ex: "VaultUpdateNotifications@company.com"
3. SMTPServer - The FQDN or IP address of the SMTP server the vault will be using to send the email notifications. Keep in mind that if FQDN is used, the Vault will need to be able to resolve the name
4. AutomaticReboot - A boolean value (default false) whether or not the server should do an automatic reboot after updates have been installed and a reboot is expected

These parameters are also set on the ServiceCheckOnReboot.ps1 script, except the AutomaticReboot one. These parameters can be inserted as a part of the command-line in scheduled tasks.
