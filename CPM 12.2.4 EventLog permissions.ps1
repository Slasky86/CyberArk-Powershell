param (

[string]$Principle = ".\ScannerUser"
)

# Get SDDL
$orgSDDL = Get-ACL ("HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\" | Select -ExpandProperty SDDL;

Write-Host "Before:"
$orgSDDL;

# Create ACL
$acl = New-Object System.Security.AccessControl.RegistrySecurity;
$acl.SetSecurityDescriptorSddlForm($orgSDDL);

# Create ACE
$ACE = New-Object System.Security.AccessControl.RegistryAccessRule $Principle,"FullControl","ContainerInherit,ObjectInherit","None","Allow"

# Combine ACL
$acl.AddAccessRule($ACE)
$newSDDL = $acl.Sddl;

Write-host "After:"
$newSDDL

# Store SDDL
Set-Acl -Path ("HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\") -AclObject $acl;

# Compose Key

$logpath = ("HKLM:\SYSTEM\CurrentControlSet\Services\EventLog\")
if (Test-Path $logpath) {

    $acl = Get-Acl $logpath
    $ace = New-Object System.Security.AccessControl.RegistryAccessRule $Principle,"FullControl","ContainerInherit,ObjectInherit","None","Allow"
    $acl.AddAccessRule($ace)
    Set-Acl $logpath $acl

}

else {

    Write-Error "Cannot access EventLog entry in registry"

}