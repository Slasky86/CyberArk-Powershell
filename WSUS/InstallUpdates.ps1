[CmdletBinding()]
param (
	[Parameter(Mandatory)]
	[string[]]$Recipient,
	[Parameter(Mandatory)]
	$SendFrom,
	[Parameter(Mandatory)]
	$SMTPServer
)


Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 5000 -Message "Vault WU installation script started"

function WriteRed($text)
{
	Write-Host $text -ForegroundColor Red
}
function WriteRedNoNewLine($text)
{
	Write-Host $text -ForegroundColor Red -NoNewLine
}
function WriteGreen($text)
{
	Write-Host $text -ForegroundColor Green
}
function DeleteFWRule($WSUSRuleName)
{
	#Delete Firewall rule
	netsh advfirewall firewall delete rule name=$WSUSRuleName dir=out | Out-Null
    netsh advfirewall firewall delete rule name=SMTP-Out dir=out | Out-Null
	""
	"[Firewall rule deleted]"
	""
}
function CloseServices($wuauservName, $TrustedInstallerName, $UpdateOrchestratorName)
{
	#Stoping & Disabling "Windows Update" services
	Set-Service -Name $wuauservName -StartupType Disabled
	Get-Service -Name $wuauservName | Stop-Service -Force 
	If (Get-Service $UpdateOrchestratorName -ErrorAction SilentlyContinue) {
		Set-Service -Name $UpdateOrchestratorName -StartupType Disabled
		Get-Service -Name $UpdateOrchestratorName | Stop-Service -Force 
	}
	Set-Service -Name $TrustedInstallerName -StartupType Disabled
	$i = 0
	$maxTries = 30
	$waitSeconds = 10
	While($i -le $maxTries)
	{
		Try
		{
			Get-Service -Name $TrustedInstallerName | Stop-Service -Force -ErrorAction Stop
			""
			"Windows update services disabled."
			"Ok."
			""
			break
		}
		Catch
		{
			if($i -eq 0)
			{
				WriteRedNoNewLine ("Waiting " + ($waitSeconds * $maxTries) + " seconds for TrustedInstaller to finish.")
			}
			else
			{
				WriteRedNoNewLine "."
			}
			if($i -eq $maxTries)
			{
				""
				WriteRed "Failed to stop TrustedInstaller"
				WriteRed "It may be related to the updates needed restart."
				WriteRed "After PC restart check this service is stopped."
				break
			}
			$i++
			Start-Sleep -s $waitSeconds
		}
	}
	""
	WriteGreen "***** Windows Updates Installer Finised *****"
}
function CheckPort($portNumber)
{
	if((([string]($portNumber)).Contains(".")) -or (-NOT [bool]($portNumber -as [int])) -or ([int]$portNumber -lt 1) -or ([int]$portNumber -gt 65535))
	{
		return 0
	}
	return 1
}

if (-NOT [string]::IsNullOrEmpty($args[0]))
{
	$InputPort1 = $args[0]
	if(-NOT (CheckPort $InputPort1))
	{
		WriteRed "First port is wrong, should be [1-65535]"
		return
	}
}
if (-NOT [string]::IsNullOrEmpty($args[1]))
{
	$InputPort2 = $args[1]
	if(-NOT (CheckPort $InputPort2))
	{
		WriteRed "Second port is wrong, should be [1-65535]"
		return
	}
}

WriteGreen "***** Windows Updates Installer *****"
""

# Variables for the registery path
$WsusRegistryPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate"

# Names of the registery variable
$WUServer = "WUServer"
$wuauservName = "wuauserv"
$TrustedInstallerName = "TrustedInstaller"
$UpdateOrchestratorName = "UsoSvc"
$WSUSRuleName = "WSUS Outbound port"

#Set services start type
"Configuring windows update services..."
Set-Service -Name $wuauservName -StartupType automatic
Set-Service -Name $TrustedInstallerName -StartupType manual
If (Get-Service $UpdateOrchestratorName -ErrorAction SilentlyContinue) {
	Set-Service -Name $UpdateOrchestratorName -StartupType manual
}
"Ok."
""

#Get the value of WSUS location from the registery
Try
{
	if(Test-Path $WsusRegistryPath)
	{
		$WSUSServer = Get-ItemProperty -Path $WsusRegistryPath | Select-Object -ExpandProperty $WUServer
	}
	else
	{
		WriteRed ($WsusRegistryPath + " - path doesn't exists")
		CloseServices $wuauservName $TrustedInstallerName $UpdateOrchestratorName
		return
	}
}
Catch
{
	WriteRed "Something is wrong, please check if WSUSServer is exists in registery"
	CloseServices $wuauservName $TrustedInstallerName $UpdateOrchestratorName
	return
}

if($WSUSServer.Contains(":"))
{
	$SplitIndex = $WSUSServer.LastIndexOf(":")
	$WSUSPort = $WSUSServer.Substring($SplitIndex + 1)
	$WSUSDNS = $WSUSServer.Substring(0, $SplitIndex)
	if($WSUSDNS.ToLower().Contains("http://"))
	{
		$WSUSDNS = $WSUSDNS.Substring(7)
	}	
	else 
	{
		if($WSUSDNS.ToLower().Contains("https://"))
		{
			$WSUSDNS = $WSUSDNS.Substring(8)
		}
	}
	Try
	{
		$WSUSIp = [System.Net.Dns]::GetHostByName($WSUSDNS).AddressList.IPAddressToString
		if(-NOT $WSUSIp)
		{
			$WSUSIp = [System.Net.Dns]::GetHostByName($WSUSDNS).HostName
		}
	}
	Catch
	{
		WriteRed "Can't find WSUS IP. In case you are using DNS record, please update hosts file"
		CloseServices $wuauservName $TrustedInstallerName $UpdateOrchestratorName
		return
	}
}
else
{
	WriteRed "Wrong WSUSServer, port not found"
	CloseServices $wuauservName $TrustedInstallerName $UpdateOrchestratorName
	return
}

"Configuring the firewall..."
$allPorts = '"'
if($InputPort1) 
{
	$allPorts = $allPorts + $InputPort1 + ','
}
if($InputPort2) 
{
	$allPorts = $allPorts + $InputPort2 + ','
}
$allPorts = $allPorts + $WSUSPort +'"'
netsh advfirewall firewall add rule name=$WSUSRuleName dir=out action=allow protocol=TCP remoteport=$allPorts remoteip=$WSUSIp
netsh advfirewall firewall add rule name=SMTP-Out dir=out action=allow protocol=TCP remoteport=25 remoteip=$SMTPServer

#-----------------------------------------------------------------------------#
###################    Installing windows updates    ##########################
#-----------------------------------------------------------------------------#
""
$Session = New-Object -ComObject Microsoft.Update.Session
"Looking for windows updates..."
$SearchStr = "IsInstalled=0"
$Searcher = $Session.CreateUpdateSearcher()
Try
{
	$SearcherResult = $Searcher.Search($SearchStr).Updates
}
Catch
{
	WriteRed ("Search for updates failed, Error: " + $_.Exception.Message + " Failed Item: " + $_.Exception.ItemName)
	DeleteFWRule $WSUSRuleName
	CloseServices $wuauservName $TrustedInstallerName $UpdateOrchestratorName
	return
}

DeleteFWRule $WSUSRuleName

if($SearcherResult.Count -eq 0)
{
	WriteGreen "No updates found"
    Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 5001 -Message "No updates found"
}
else
{
	$downloadCount = 0
	For ($i=0; $i -lt $SearcherResult.Count; $i++)
	{
		if($SearcherResult.Item($i).IsDownloaded)
		{
			$downloadCount++
		}
	}
	
	"Found " + $SearcherResult.Count + " Updates"
	"" + $downloadCount + " Updates already downloaded and will be installed:"
	
	For ($i=0; $i -lt $SearcherResult.Count; $i++)
	{
		if($SearcherResult.Item($i).IsDownloaded)
		{
			"" + ($i + 1) + ": " + $SearcherResult.Item($i).Title
		}
	}

    Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 5002 -Message ("Found " + $SearcherResult.Count + " Updates. " + $downloadCount + " updates already downloaded and will be installed.")
	
	WriteGreen ("" + $downloadCount + " Updates already downloaded and will be installed:")
	""
	
	$NumberOfUpdate = 0
	$NotInstalledCount = 0
	$ErrorCount = 0
	$NeedsReboot = $false
	"Installing " + $downloadCount + " Updates:"
	For ($i=0; $i -lt $SearcherResult.Count; $i++)
	{
		$Update = $SearcherResult.Item($i)
		if($Update.IsDownloaded)
		{
			$NumberOfUpdate++
			"Installing update: " + $Update.Title
			$objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
			$objCollectionTmp.Add($Update) | Out-Null
			
			$objInstaller = $Session.CreateUpdateInstaller()
			$objInstaller.Updates = $objCollectionTmp
			Try
			{
				$InstallResult = $objInstaller.Install()
			}
			Catch
			{
				$ErrorCount++
				If($_ -match "HRESULT: 0x80240044")
				{
					WriteRed "Your security policy don't allow a non-administator identity to perform this task"
				}
				else
				{
					WriteRed $_
				}
				continue
			}
			
			If(!$NeedsReboot) 
			{ 
				$NeedsReboot = $installResult.RebootRequired
			} 
			
			Switch -exact ($InstallResult.ResultCode)
			{
				0   { $Status = "NotStarted"}
				1   { $Status = "InProgress"}
				2   { $Status = "Installed"}
				3   { $Status = "InstalledWithErrors"}
				4   { $Status = "Failed"}
				5   { $Status = "Aborted"}
			}
				   
			Switch($Update.MaxDownloadSize)
			{
				{[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
				{[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }  
				{[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }    
				{[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
				default { $size = $_+"B" }
			}
			
			$log = New-Object PSObject -Property @{
				Title = $Update.Title
				Number = "" + $NumberOfUpdate + "/" + $downloadCount
				KB = "KB" + $Update.KBArticleIDs
				Size = $size
				Status = $Status
			}
			
			if(-NOT $Status -eq "Installed")
			{
				$NotInstalledCount++
				WriteRed $Status
			}
			$log
		}
	}
	WriteGreen ("" + ($NumberOfUpdate - $NotInstalledCount - $ErrorCount) + " updates installed")
    Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 5003 -Message (($NumberOfUpdate - $NotInstalledCount - $ErrorCount) + " updates installed")
	if($NotInstalledCount -gt 0)
	{
		WriteRed ("" + $NotInstalledCount + " updates not installed")
        Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType error -EventId 5005 -Message ($NotInstalledCount + " updates not installed")
	}
	if($ErrorCount -gt 0)
	{
		WriteRed ("" + $ErrorCount + " updates with errors")
        Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType error -EventId 5005 -Message ($ErrorCount + " updates with errors")
	}
	If ($NeedsReboot)
	{ 
		WriteRed "Computer reboot is needed"
        Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 5090 -Message "Reboot needed. The server will automatically reboot"
	}
}
""
CloseServices $wuauservName $TrustedInstallerName $UpdateOrchestratorName

$installedupdates = $NumberOfUpdate - $NotInstalledCount - $ErrorCount

if ($installedupdates -in 0, $null ,"") {

    $mailbody = "No Windows Updates installed on $env:computername."

}

if ($NumberOfUpdate -in $null, "") {

    $NumberOfUpdate = 0

}

if ($ErrorCount -in $null, "") {

    $ErrorCount = 0

}

$mailbody = "Windows Updates installed on $env:computername. `n`n`n$installedupdates updates was installed `n$ErrorCount updates had errors `n$NotInstalledCount updates were not installed"

Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Windows Update' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer



If ($NeedsReboot) {
    
    Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 5099 -Message "Windows update script finished. A reboot is required, rebooting now!"
    Return $NeedsReboot


}

else {

    Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 5099 -Message "Windows update script finished."

}


# SIG # Begin signature block
# MIIfdQYJKoZIhvcNAQcCoIIfZjCCH2ICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCARFS+KrdD6GoJF
# Tf+4lVJJoEW8zocvIMRZXdIzybzzPqCCDnUwggROMIIDNqADAgECAg0B7l8Wnf+X
# NStkZdZqMA0GCSqGSIb3DQEBCwUAMFcxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBH
# bG9iYWxTaWduIG52LXNhMRAwDgYDVQQLEwdSb290IENBMRswGQYDVQQDExJHbG9i
# YWxTaWduIFJvb3QgQ0EwHhcNMTgwOTE5MDAwMDAwWhcNMjgwMTI4MTIwMDAwWjBM
# MSAwHgYDVQQLExdHbG9iYWxTaWduIFJvb3QgQ0EgLSBSMzETMBEGA1UEChMKR2xv
# YmFsU2lnbjETMBEGA1UEAxMKR2xvYmFsU2lnbjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAMwldpB5BngiFvXAg7aEyiie/QV2EcWtiHL8RgJDx7KKnQRf
# JMsuS+FggkbhUqsMgUdwbN1k0ev1LKMPgj0MK66X17YUhhB5uzsTgHeMCOFJ0mpi
# Lx9e+pZo34knlTifBtc+ycsmWQ1z3rDI6SYOgxXG71uL0gRgykmmKPZpO/bLyCiR
# 5Z2KYVc3rHQU3HTgOu5yLy6c+9C7v/U9AOEGM+iCK65TpjoWc4zdQQ4gOsC0p6Hp
# sk+QLjJg6VfLuQSSaGjlOCZgdbKfd/+RFO+uIEn8rUAVSNECMWEZXriX7613t2Sa
# er9fwRPvm2L7DWzgVGkWqQPabumDk3F2xmmFghcCAwEAAaOCASIwggEeMA4GA1Ud
# DwEB/wQEAwIBBjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBSP8Et/qC5FJK5N
# UPpjmove4t0bvDAfBgNVHSMEGDAWgBRge2YaRQ2XyolQL30EzTSo//z9SzA9Bggr
# BgEFBQcBAQQxMC8wLQYIKwYBBQUHMAGGIWh0dHA6Ly9vY3NwLmdsb2JhbHNpZ24u
# Y29tL3Jvb3RyMTAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3JsLmdsb2JhbHNp
# Z24uY29tL3Jvb3QuY3JsMEcGA1UdIARAMD4wPAYEVR0gADA0MDIGCCsGAQUFBwIB
# FiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzANBgkqhkiG
# 9w0BAQsFAAOCAQEAI3Dpz+K+9VmulEJvxEMzqs0/OrlkF/JiBktI8UCIBheh/qvR
# XzzGM/Lzjt0fHT7MGmCZggusx/x+mocqpX0PplfurDtqhdbevUBj+K2myIiwEvz2
# Qd8PCZceOOpTn74F9D7q059QEna+CYvCC0h9Hi5R9o1T06sfQBuKju19+095VnBf
# DNOOG7OncA03K5eVq9rgEmscQM7Fx37twmJY7HftcyLCivWGQ4it6hNu/dj+Qi+5
# fV6tGO+UkMo9J6smlJl1x8vTe/fKTNOvUSGSW4R9K58VP3TLUeiegw4WbxvnRs4j
# vfnkoovSOWuqeRyRLOJhJC2OKkhwkMQexejgcDCCBKcwggOPoAMCAQICDkgbagep
# Qkweqv7zzfEPMA0GCSqGSIb3DQEBCwUAMEwxIDAeBgNVBAsTF0dsb2JhbFNpZ24g
# Um9vdCBDQSAtIFIzMRMwEQYDVQQKEwpHbG9iYWxTaWduMRMwEQYDVQQDEwpHbG9i
# YWxTaWduMB4XDTE2MDYxNTAwMDAwMFoXDTI0MDYxNTAwMDAwMFowbjELMAkGA1UE
# BhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExRDBCBgNVBAMTO0dsb2Jh
# bFNpZ24gRXh0ZW5kZWQgVmFsaWRhdGlvbiBDb2RlU2lnbmluZyBDQSAtIFNIQTI1
# NiAtIEczMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA2be6Ja2U81u+
# QQYcU8oMEIxRQVkzeWT0V53k1SXE7FCEWJhyeUDiL3jUkuomDp6ulXz7xP1xRN2M
# X7cji1679PxLyyM9w3YD9dGMRbxxdR2L0omJvuNRPcbIirIxNQduufW6ag30EJ+u
# 1WJJKHvsV7qrMnyxfdKiVgY27rDv0Gqu6qsf1g2ffJb7rXCZLV2V8IDQeUbsVTrM
# 0zj7BAeoB3WCguDQfne4j+vSKPyubRRoQX92Q9dIumBE4bdy6NDwIAN72tq0BnXH
# sgPe+JTGaI9ee56bnTbgztJrxsZr6RQitXF+to9aH9vnbvRCEJBo5itFEE9zuizX
# xTFqct1jcwIDAQABo4IBYzCCAV8wDgYDVR0PAQH/BAQDAgEGMB0GA1UdJQQWMBQG
# CCsGAQUFBwMDBggrBgEFBQcDCTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQW
# BBTcLFgsKm81LZ95lahIXcRtPlO/uTAfBgNVHSMEGDAWgBSP8Et/qC5FJK5NUPpj
# move4t0bvDA+BggrBgEFBQcBAQQyMDAwLgYIKwYBBQUHMAGGImh0dHA6Ly9vY3Nw
# Mi5nbG9iYWxzaWduLmNvbS9yb290cjMwNgYDVR0fBC8wLTAroCmgJ4YlaHR0cDov
# L2NybC5nbG9iYWxzaWduLmNvbS9yb290LXIzLmNybDBiBgNVHSAEWzBZMAsGCSsG
# AQQBoDIBAjAHBgVngQwBAzBBBgkrBgEEAaAyAV8wNDAyBggrBgEFBQcCARYmaHR0
# cHM6Ly93d3cuZ2xvYmFsc2lnbi5jb20vcmVwb3NpdG9yeS8wDQYJKoZIhvcNAQEL
# BQADggEBAHYJxMwv2e8eS6n4V/NAOSHKTDwdnikrINQrRNKIzhoNBc+Dgbvrabwx
# jSrEx0TMYGCUHM+h4QIkDq1bvizCJx5nt+goHzJR4znzmN+4ny6LKrR7CgO8vTYE
# j8nQnE+jAieZsPBF6TTf5DqjtwY32G8qeZDU1E5YcexTqWGY9zlp4BKcV1hyhicp
# pR3lMvMrmZdavyuwPLQG6g5k7LfNZYAkF8LZN/WxJhA1R3uaArpUokWT/3m/GozF
# n7Wf33jna1DxR5RpSyS42gXoDJ1PBuxKMSB+T12GhC81o82cwYRXHx+twOKkse8p
# ayGXptT+7QM3sPz1jSq83ISD497D518wggV0MIIEXKADAgECAgwhXYQh+9kPSKH6
# QS4wDQYJKoZIhvcNAQELBQAwbjELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2Jh
# bFNpZ24gbnYtc2ExRDBCBgNVBAMTO0dsb2JhbFNpZ24gRXh0ZW5kZWQgVmFsaWRh
# dGlvbiBDb2RlU2lnbmluZyBDQSAtIFNIQTI1NiAtIEczMB4XDTE5MDQwMjE0MDI0
# NVoXDTIyMDQwMjE0MDI0NVowgcgxHTAbBgNVBA8MFFByaXZhdGUgT3JnYW5pemF0
# aW9uMRIwEAYDVQQFEwk1MTIyOTE2NDIxEzARBgsrBgEEAYI3PAIBAxMCSUwxCzAJ
# BgNVBAYTAklMMRkwFwYDVQQIExBDZW50cmFsIERpc3RyaWN0MRQwEgYDVQQHEwtQ
# ZXRhaCBUaWt2YTEfMB0GA1UEChMWQ3liZXJBcmsgU29mdHdhcmUgTHRkLjEfMB0G
# A1UEAxMWQ3liZXJBcmsgU29mdHdhcmUgTHRkLjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAJmp1fuFtNzvXmXAG4MZy5nl5gLRMycA6ieFpbOIPdMOTMvO
# wWaW4VASvtzqyZOpUNV0OZka6ajkVrM7IzihX43zvfEizWmG+359QU6htgHSWmII
# KDjEOxQrnq/+l0qgbBge6zqA4mzXh+frgpgnfvL9Rq7WTCjNywTl7UD3mn5VuKbZ
# XIhn19ICv7WKSr/VVoGNpIy/o3PmgHLfSMX9vUaxU+sXIZKhP1eqFtMMllO0jzK2
# hAttOAGLlKJO2Yp17+HOI86vfVAJ8YGOeFdtObgdrL/DhSORMFZE5Y5eT14vLZQu
# OODTz/YZE/PnrwxGKFqPQNHo9O7/j4kNxGTa1m8CAwEAAaOCAbUwggGxMA4GA1Ud
# DwEB/wQEAwIHgDCBoAYIKwYBBQUHAQEEgZMwgZAwTgYIKwYBBQUHMAKGQmh0dHA6
# Ly9zZWN1cmUuZ2xvYmFsc2lnbi5jb20vY2FjZXJ0L2dzZXh0ZW5kY29kZXNpZ25z
# aGEyZzNvY3NwLmNydDA+BggrBgEFBQcwAYYyaHR0cDovL29jc3AyLmdsb2JhbHNp
# Z24uY29tL2dzZXh0ZW5kY29kZXNpZ25zaGEyZzMwVQYDVR0gBE4wTDBBBgkrBgEE
# AaAyAQIwNDAyBggrBgEFBQcCARYmaHR0cHM6Ly93d3cuZ2xvYmFsc2lnbi5jb20v
# cmVwb3NpdG9yeS8wBwYFZ4EMAQMwCQYDVR0TBAIwADBFBgNVHR8EPjA8MDqgOKA2
# hjRodHRwOi8vY3JsLmdsb2JhbHNpZ24uY29tL2dzZXh0ZW5kY29kZXNpZ25zaGEy
# ZzMuY3JsMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQQP3rH7GUJCWmd
# tvKh9RqkZNQaEjAfBgNVHSMEGDAWgBTcLFgsKm81LZ95lahIXcRtPlO/uTANBgkq
# hkiG9w0BAQsFAAOCAQEAtRWdBsZ830FMJ9GxODIHyFS0z08inqP9c3iNxDk3BYNL
# WxtU91cGtFdnCAc8G7dNMEQ+q0TtQKTcJ+17k6GdNM8Lkanr51MngNOl8CP6QMr+
# rIzKAipex1J61Mf44/6Y6gOMGHW7jk84QxMSEbYIglfkHu+RhH8mhYRGKGgHOX3R
# ViIoIxthvlG08/nTux3zeVnSAmXB5Z8KJ+FTzLyZhFii2i2TLAt/a95dMOb4YquH
# qK9lmeFCLovYNIAihC7NHBruSGkt/sguM/17JWPpgHpjJxrIZH3dVH41LNPb3Bz2
# KDHmv37ZRpQvuxAyctrTAPA6HJtuEJnIo6DhFR9LfTGCEFYwghBSAgEBMH4wbjEL
# MAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExRDBCBgNVBAMT
# O0dsb2JhbFNpZ24gRXh0ZW5kZWQgVmFsaWRhdGlvbiBDb2RlU2lnbmluZyBDQSAt
# IFNIQTI1NiAtIEczAgwhXYQh+9kPSKH6QS4wDQYJYIZIAWUDBAIBBQCgfDAQBgor
# BgEEAYI3AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgXY4yYqKyiTJk
# 0NwgSXqAZu2NbxIPgy3E0SE7WkY/0QIwDQYJKoZIhvcNAQEBBQAEggEAEgsgUVZD
# qZHwFfUOS9O5PtVWcXpL5o6UzwiN6AdNS98ycNJ0636TSUq6LHd4NmOVjuCxDhMy
# nrFjlG/t80Qsc/jkCH3gKTY+TX8Q441k7KqOjTs9X2vUDvF/Or9dRWyhtx+jJggb
# wBZeoPdaBUt3bfRBfQe4s8kIBTLlS5j2/G8ZaNZ3mhr4VnQEa6d0bNUs9spmAs9k
# 86uDA3bPBtWO7EEgscEhL67Zxsi2L8NeJmYni7rva7Z5Vi2dWBJz+ZLSvW+wqo7X
# Quif61S+yUnl6N/Pip9XhkGgvoZsQDlQF/XSFRNr1nJDOS4jn1WvQLvv1G6ow1ZN
# tPfvv50xQhnbh6GCDiswgg4nBgorBgEEAYI3AwMBMYIOFzCCDhMGCSqGSIb3DQEH
# AqCCDgQwgg4AAgEDMQ0wCwYJYIZIAWUDBAIBMIH+BgsqhkiG9w0BCRABBKCB7gSB
# 6zCB6AIBAQYLYIZIAYb4RQEHFwMwITAJBgUrDgMCGgUABBQPK6iEKBKSu2kLN+5d
# ubNOxOn/fwIUDxK5PhuqcrQO9EJyNdWgu5BlaOoYDzIwMjAwNzIwMTc1NDQwWjAD
# AgEeoIGGpIGDMIGAMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29y
# cG9yYXRpb24xHzAdBgNVBAsTFlN5bWFudGVjIFRydXN0IE5ldHdvcmsxMTAvBgNV
# BAMTKFN5bWFudGVjIFNIQTI1NiBUaW1lU3RhbXBpbmcgU2lnbmVyIC0gRzOgggqL
# MIIFODCCBCCgAwIBAgIQewWx1EloUUT3yYnSnBmdEjANBgkqhkiG9w0BAQsFADCB
# vTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDlZlcmlTaWduLCBJbmMuMR8wHQYDVQQL
# ExZWZXJpU2lnbiBUcnVzdCBOZXR3b3JrMTowOAYDVQQLEzEoYykgMjAwOCBWZXJp
# U2lnbiwgSW5jLiAtIEZvciBhdXRob3JpemVkIHVzZSBvbmx5MTgwNgYDVQQDEy9W
# ZXJpU2lnbiBVbml2ZXJzYWwgUm9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTAe
# Fw0xNjAxMTIwMDAwMDBaFw0zMTAxMTEyMzU5NTlaMHcxCzAJBgNVBAYTAlVTMR0w
# GwYDVQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjEfMB0GA1UECxMWU3ltYW50ZWMg
# VHJ1c3QgTmV0d29yazEoMCYGA1UEAxMfU3ltYW50ZWMgU0hBMjU2IFRpbWVTdGFt
# cGluZyBDQTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALtZnVlVT52M
# cl0agaLrVfOwAa08cawyjwVrhponADKXak3JZBRLKbvC2Sm5Luxjs+HPPwtWkPhi
# G37rpgfi3n9ebUA41JEG50F8eRzLy60bv9iVkfPw7mz4rZY5Ln/BJ7h4OcWEpe3t
# r4eOzo3HberSmLU6Hx45ncP0mqj0hOHE0XxxxgYptD/kgw0mw3sIPk35CrczSf/K
# O9T1sptL4YiZGvXA6TMU1t/HgNuR7v68kldyd/TNqMz+CfWTN76ViGrF3PSxS9TO
# 6AmRX7WEeTWKeKwZMo8jwTJBG1kOqT6xzPnWK++32OTVHW0ROpL2k8mc40juu1MO
# 1DaXhnjFoTcCAwEAAaOCAXcwggFzMA4GA1UdDwEB/wQEAwIBBjASBgNVHRMBAf8E
# CDAGAQH/AgEAMGYGA1UdIARfMF0wWwYLYIZIAYb4RQEHFwMwTDAjBggrBgEFBQcC
# ARYXaHR0cHM6Ly9kLnN5bWNiLmNvbS9jcHMwJQYIKwYBBQUHAgIwGRoXaHR0cHM6
# Ly9kLnN5bWNiLmNvbS9ycGEwLgYIKwYBBQUHAQEEIjAgMB4GCCsGAQUFBzABhhJo
# dHRwOi8vcy5zeW1jZC5jb20wNgYDVR0fBC8wLTAroCmgJ4YlaHR0cDovL3Muc3lt
# Y2IuY29tL3VuaXZlcnNhbC1yb290LmNybDATBgNVHSUEDDAKBggrBgEFBQcDCDAo
# BgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1wLTIwNDgtMzAdBgNVHQ4E
# FgQUr2PWyqNOhXLgp7xB8ymiOH+AdWIwHwYDVR0jBBgwFoAUtnf6aUhHn1MS1cLq
# BzJ2B9GXBxkwDQYJKoZIhvcNAQELBQADggEBAHXqsC3VNBlcMkX+DuHUT6Z4wW/X
# 6t3cT/OhyIGI96ePFeZAKa3mXfSi2VZkhHEwKt0eYRdmIFYGmBmNXXHy+Je8Cf0c
# kUfJ4uiNA/vMkC/WCmxOM+zWtJPITJBjSDlAIcTd1m6JmDy1mJfoqQa3CcmPU1dB
# kC/hHk1O3MoQeGxCbvC2xfhhXFL1TvZrjfdKer7zzf0D19n2A6gP41P3CnXsxnUu
# qmaFBJm3+AZX4cYO9uiv2uybGB+queM6AL/OipTLAduexzi7D1Kr0eOUA2AKTaD+
# J20UMvw/l0Dhv5mJ2+Q5FL3a5NPD6itas5VYVQR9x5rsIwONhSrS/66pYYEwggVL
# MIIEM6ADAgECAhB71OWvuswHP6EBIwQiQU0SMA0GCSqGSIb3DQEBCwUAMHcxCzAJ
# BgNVBAYTAlVTMR0wGwYDVQQKExRTeW1hbnRlYyBDb3Jwb3JhdGlvbjEfMB0GA1UE
# CxMWU3ltYW50ZWMgVHJ1c3QgTmV0d29yazEoMCYGA1UEAxMfU3ltYW50ZWMgU0hB
# MjU2IFRpbWVTdGFtcGluZyBDQTAeFw0xNzEyMjMwMDAwMDBaFw0yOTAzMjIyMzU5
# NTlaMIGAMQswCQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRp
# b24xHzAdBgNVBAsTFlN5bWFudGVjIFRydXN0IE5ldHdvcmsxMTAvBgNVBAMTKFN5
# bWFudGVjIFNIQTI1NiBUaW1lU3RhbXBpbmcgU2lnbmVyIC0gRzMwggEiMA0GCSqG
# SIb3DQEBAQUAA4IBDwAwggEKAoIBAQCvDoqq+Ny/aXtUF3FHCb2NPIH4dBV3Z5Cc
# /d5OAp5LdvblNj5l1SQgbTD53R2D6T8nSjNObRaK5I1AjSKqvqcLG9IHtjy1GiQo
# +BtyUT3ICYgmCDr5+kMjdUdwDLNfW48IHXJIV2VNrwI8QPf03TI4kz/lLKbzWSPL
# gN4TTfkQyaoKGGxVYVfR8QIsxLWr8mwj0p8NDxlsrYViaf1OhcGKUjGrW9jJdFLj
# V2wiv1V/b8oGqz9KtyJ2ZezsNvKWlYEmLP27mKoBONOvJUCbCVPwKVeFWF7qhUhB
# IYfl3rTTJrJ7QFNYeY5SMQZNlANFxM48A+y3API6IsW0b+XvsIqbAgMBAAGjggHH
# MIIBwzAMBgNVHRMBAf8EAjAAMGYGA1UdIARfMF0wWwYLYIZIAYb4RQEHFwMwTDAj
# BggrBgEFBQcCARYXaHR0cHM6Ly9kLnN5bWNiLmNvbS9jcHMwJQYIKwYBBQUHAgIw
# GRoXaHR0cHM6Ly9kLnN5bWNiLmNvbS9ycGEwQAYDVR0fBDkwNzA1oDOgMYYvaHR0
# cDovL3RzLWNybC53cy5zeW1hbnRlYy5jb20vc2hhMjU2LXRzcy1jYS5jcmwwFgYD
# VR0lAQH/BAwwCgYIKwYBBQUHAwgwDgYDVR0PAQH/BAQDAgeAMHcGCCsGAQUFBwEB
# BGswaTAqBggrBgEFBQcwAYYeaHR0cDovL3RzLW9jc3Aud3Muc3ltYW50ZWMuY29t
# MDsGCCsGAQUFBzAChi9odHRwOi8vdHMtYWlhLndzLnN5bWFudGVjLmNvbS9zaGEy
# NTYtdHNzLWNhLmNlcjAoBgNVHREEITAfpB0wGzEZMBcGA1UEAxMQVGltZVN0YW1w
# LTIwNDgtNjAdBgNVHQ4EFgQUpRMBqZ+FzBtuFh5fOzGqeTYAex0wHwYDVR0jBBgw
# FoAUr2PWyqNOhXLgp7xB8ymiOH+AdWIwDQYJKoZIhvcNAQELBQADggEBAEaer/C4
# ol+imUjPqCdLIc2yuaZycGMv41UpezlGTud+ZQZYi7xXipINCNgQujYk+gp7+zvT
# Yr9KlBXmgtuKVG3/KP5nz3E/5jMJ2aJZEPQeSv5lzN7Ua+NSKXUASiulzMub6KlN
# 97QXWZJBw7c/hub2wH9EPEZcF1rjpDvVaSbVIX3hgGd+Yqy3Ti4VmuWcI69bEepx
# qUH5DXk4qaENz7Sx2j6aescixXTN30cJhsT8kSWyG5bphQjo3ep0YG5gpVZ6DchE
# WNzm+UgUnuW/3gC9d7GYFHIUJN/HESwfAD/DSxTGZxzMHgajkF9cVIs+4zNbgg/F
# t4YCTnGf6WZFP3YxggJaMIICVgIBATCBizB3MQswCQYDVQQGEwJVUzEdMBsGA1UE
# ChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xHzAdBgNVBAsTFlN5bWFudGVjIFRydXN0
# IE5ldHdvcmsxKDAmBgNVBAMTH1N5bWFudGVjIFNIQTI1NiBUaW1lU3RhbXBpbmcg
# Q0ECEHvU5a+6zAc/oQEjBCJBTRIwCwYJYIZIAWUDBAIBoIGkMBoGCSqGSIb3DQEJ
# AzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjAwNzIwMTc1NDQwWjAv
# BgkqhkiG9w0BCQQxIgQgssH/X0E9bQPslc+ihaymlDBA2yiFg7REXQcq8wsPZWcw
# NwYLKoZIhvcNAQkQAi8xKDAmMCQwIgQgxHTOdgB9AjlODaXk3nwUxoD54oIBPP72
# U+9dtx/fYfgwCwYJKoZIhvcNAQEBBIIBABv9e2+X9gNJBDfQnSsO3U/Nr5Eu2PCi
# ZrvfcP0xj6eTWij9bIdJ48/aZSwBrSO8bPxMyjxhxYkciFpAID4NX6BVJgH99wag
# rY6/qDsoY8ZOQNBGqwDHNCeQkx5GwkLpB18FzVEIMqNHstttbFVp1BgKVdkYHudS
# NUCwVGydGN2W7WWBzscldXlJ5futLxVkvrQx5Ak+gPTo7j24dFpjwMISOuKL6wme
# /jq10PbUAh1C4aIy1HcBh8nahaTAq/Q0NqATiqvbXmWK/U5BhCZbhP9z13yqrjYu
# KaVUAlfH232GIOctc9jWVNkRRrDhJJCkr/rxcHJ+2NMEjJcxfDwBUoA=
# SIG # End signature block
