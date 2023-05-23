param (

    [string]$Principle = ".\ScannerUser",
    [string[]]$LogNames = ('Security','Application','System')

)


foreach ($LogName in $LogNames) {

    # Get SDDL
    $orgSDDL = Get-ACL ("HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\"+$LogName) | Select -exp SDDL

    Write-Host 'Before:'
    $orgSDDL

    # Create ACL
    $acl = New-Object System.Security.AccessControl.RegistrySecurity
    $acl.SetSecurityDescriptorSddlForm($orgSDDL)

    # Create ACE
    $ACE = New-Object System.Security.AccessControl.RegistryAccessRule $Principle,"FullControl","ContainerInherit,ObjectInherit","None","Allow"

    # Combine ACL
    $acl.AddAccessRule($ACE)
    $newSDDL = $acl.Sddl

    Write-Host "After:"
    $newSDDL

    #Store SDDL
    Set-Acl -Path ("HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\"+$LogName) -AclObject $acl

    #Compose Key
    $LogPath = ("HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\"+$LogName) 
    if (Test-Path $LogPath) {

        $acl = Get-Acl $LogPath
        $ace = New-Object System.Security.AccessControl.RegistryAccessRule $Principle,"FullControl","ContainerInherit,ObjectInherit","None","Allow"
        $acl.AddAccessRule($ACE)
        Set-Acl $LogPath $acl

    }

    else {

        Write-Error "Cannot access log $LogName"

    }
}