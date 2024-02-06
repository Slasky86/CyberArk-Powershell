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

try {

    Import-Module psPAS

}

catch {

    Write-Host -ForegroundColor Yellow "Powershell module psPAS not found, trying to install"
    Install-Module -name psPAS -Force

}

New-PASSession -Credential $Credential -BaseURI $PVWAURL -concurrentSession $true


foreach ($Safe in $Safes.Safename) {

    Get-PASSafe -search $Safe | Set-PASSafe -NumberOfDaysRetention $DaysOfRetention

}
