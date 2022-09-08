[CmdletBinding()]
param (
	[Parameter(Mandatory)]
	[string[]]$Recipient,
	[Parameter(Mandatory)]
	$SendFrom,
	[Parameter(Mandatory)]
	$SMTPServer
)

Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 4000 -Message "Download-script started"

function WriteRed($text)
{
	Write-Host $text -ForegroundColor Red
}
function WriteGreen($text)
{
	Write-Host $text -ForegroundColor Green
}
function TerminateEnv($WSUSRuleName, $wuauservName, $TrustedInstallerName, $UpdateOrchestratorName)
{
	#Delete Firewall rule
	netsh advfirewall firewall delete rule name=$WSUSRuleName dir=out | Out-Null
	netsh advfirewall firewall delete rule name=SMTP-Out dir=out | Out-Null
	"Firewall rule deleted."
	"Ok."
	""

	#Stoping & Disabling "Windows Update" services
	Get-Service -Name $wuauservName | Stop-Service -Force 
	Set-Service -Name $wuauservName -StartupType Disabled
	If (Get-Service $UpdateOrchestratorName -ErrorAction SilentlyContinue) {
		Get-Service -Name $UpdateOrchestratorName | Stop-Service -Force 
		Set-Service -Name $UpdateOrchestratorName -StartupType Disabled
	}
	Try
	{
		Get-Service -Name $TrustedInstallerName | Stop-Service -Force -ErrorAction Stop
	}
	Catch
	{
		WriteRed "Failed to stop TrustedInstaller"
		WriteRed "Try to stop it manually or with the 'ClosingServices' script"
		WriteRed "If fail, reboot PC and check this service is stopped."
		""
	}
	Set-Service -Name $TrustedInstallerName -StartupType Disabled
	"Windows update services disabled."
	"Ok."
	""
	WriteGreen "***** Windows Updates Downloader Finised *****"
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

WriteGreen "***** Windows Updates Downloader *****"
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
		TerminateEnv $WSUSRuleName $wuauservName $TrustedInstallerName $UpdateOrchestratorName
		return
	}
}
Catch
{
	WriteRed "Something is wrong, please check if WSUSServer is exists in registery"
	TerminateEnv $WSUSRuleName $wuauservName $TrustedInstallerName $UpdateOrchestratorName
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
		TerminateEnv $WSUSRuleName $wuauservName $TrustedInstallerName $UpdateOrchestratorName
		return
	}
}
else
{
	WriteRed "Wrong WSUSServer, port not found"
	TerminateEnv $WSUSRuleName $wuauservName $TrustedInstallerName $UpdateOrchestratorName
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

#Downloading windows updates
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
	TerminateEnv $WSUSRuleName $wuauservName $TrustedInstallerName $UpdateOrchestratorName
	return
}
if($SearcherResult.Count -eq 0)
{
	WriteGreen "No updates found"
    Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 4001 -Message "No updates found"
}
else
{
	"Found " + $SearcherResult.Count + " Updates:"
	For ($i=0; $i -lt $SearcherResult.Count; $i++)
	{
		"" + ($i + 1) + ": " + $SearcherResult.Item($i).Title
	}
	WriteGreen ("Found " + $SearcherResult.Count + " Updates")
	""
    Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information  -EventId 4002 -Message ("Found " + $SearcherResult.Count + " Updates")
	"Downloading windows updates..."
	$Downloader = $Session.CreateUpdateDownloader()
	$Downloader.Updates = $SearcherResult
	$DownloadResult = $Downloader.Download()
	$DFailed = 1
	
	Switch($DownloadResult.ResultCode)
	{
		0 { WriteRed ("Download failed, Error: NotStarted, HResult: " + $DownloadResult.HResult) 
            Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Error -EventId 4003 -Message ("Download failed, Error: NotStarted, HResult: " + $DownloadResult.HResult)}
		1 { WriteRed ("Download failed, Error: InProgress, HResult: " + $DownloadResult.HResult) 
            Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Error -EventId 4003 -Message ("Download failed, Error: InProgress, HResult: " + $DownloadResult.HResult)}
		2 
		{ 
			if($DownloadResult.HResult -eq 0)
			{
				$DFailed = 0
				WriteGreen "Download completed successfully"
                Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 4005 -Message "Download completed successfully"
			}
			else
			{
				WriteRed ("Download complete with HResult: " + $DownloadResult.HResult)
                Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Error -Message -EventId 4006 ("Download complete with HResult: " + $DownloadResult.HResult)
			}
		}
		3 { WriteRed ("Download complete with errors, HResult: " + $DownloadResult.HResult) 
            Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Error -EventId 4003 -Message  ("Download complete with errors, HResult: " + $DownloadResult.HResult) }
		4 { WriteRed ("Download failed, Error: Failed, HResult: " + $DownloadResult.HResult) 
            Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Error -EventId 4003 -Message ("Download failed, Error: Failed, HResult: " + $DownloadResult.HResult) }
		5 { WriteRed ("Download failed, Error: Aborted, HResult: " + $DownloadResult.HResult) 
            Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Error -EventId 4003 -Message ("Download failed, Error: Aborted, HResult: " + $DownloadResult.HResult)}
		default { WriteRed "Download failed with unknown error" 
                  Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Error -EventId 4003 -Message "Download failed with unknown error"}
	}
	if($DFailed)
	{
		"List of the updates failed to download:"
		for ($i=0; $i -lt $Downloader.Updates.Count; $i++)
		{
			if(-Not $Downloader.Updates.Item($i).IsDownloaded)
			{
				"" + ($i + 1) +": " + $Downloader.Updates.Item($i).Title
				$mailbody += "$($Downloader.Updates.Item($i).Title) has failed to download on $env:Computername"
			}
		}
		Send-MailMessage -From $SendFrom -To $Recipient -Subject 'CyberArk Vault Windows Update' -Body $mailbody -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer $SMTPServer
	}
}

Write-EventLog -LogName Application -Source "VaultWUUpdate" -EntryType Information -EventId 4099 -Message "Download-script completed"
""
TerminateEnv $WSUSRuleName $wuauservName $TrustedInstallerName $UpdateOrchestratorName
# SIG # Begin signature block
# MIIfdQYJKoZIhvcNAQcCoIIfZjCCH2ICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCARmCOpPnP7lt4B
# w3/n52ovtrFWZaHSODIO1me2EJJDSqCCDnUwggROMIIDNqADAgECAg0B7l8Wnf+X
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgbECTbAbH9hOw
# XY6DLVlBCLXO0rAgBjzpL87eNM+yb6owDQYJKoZIhvcNAQEBBQAEggEARoAK3vGK
# aUwnUBs76nbyWx+gl4nYHoJcWHWhCGv6p1q6+OQwp8vXR7UG+YMy8ZDrZ5gbAEPB
# kr16vPAjqgfkZBikoetFV18W4YY5TN/fApsOuy6UXZ4AYL2JN4Pw7g2MZ6FRbqdq
# GXcSxpJbDWANj1k1EZOpVLbwB4TAZiOC/XIuBOSrp7v0p9cwuLszFzt4njwegxu6
# 67LRZ0xLeRC2yajrBKGXS6dg2PeCIXmUVVm/ITaoNq2frHaGmk+98Uh5JoeKoD0f
# urOgp8aZ568C9/D3ow433WblU0K9WSv3vqabErrsiAe2HTJKOJ1e5Lt6lSgQgi0A
# DN4nw70xxi+35aGCDiswgg4nBgorBgEEAYI3AwMBMYIOFzCCDhMGCSqGSIb3DQEH
# AqCCDgQwgg4AAgEDMQ0wCwYJYIZIAWUDBAIBMIH+BgsqhkiG9w0BCRABBKCB7gSB
# 6zCB6AIBAQYLYIZIAYb4RQEHFwMwITAJBgUrDgMCGgUABBRNeF0HYq3BeHvJKdDT
# 6VYUBWrW8QIUFKglayt1mOPVyZvITajfZZ6IXxYYDzIwMjAwNzIwMTc1NDM5WjAD
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
# AzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjAwNzIwMTc1NDM5WjAv
# BgkqhkiG9w0BCQQxIgQgQXQUKEDc64TvMdPI2IiHtDurQnTMGh4wiefQseGuIn8w
# NwYLKoZIhvcNAQkQAi8xKDAmMCQwIgQgxHTOdgB9AjlODaXk3nwUxoD54oIBPP72
# U+9dtx/fYfgwCwYJKoZIhvcNAQEBBIIBAG/Vc1zcvvvz5K969CeXWAK/s3YzPX1e
# /WAVcrnjRfPnO61Y4XXodTmgJ/3rbJnEEGhUrAO79pmBwTG2AYTx4JbUdjIX85fD
# jXgF0rpCycpRnB2NRGit/FAYHAz6+cdg4L+9NqQN4kGFYSoa6wnqbqw4Ve9KsMMf
# tg4R3+xYDprCxUg6xN1NPN2kWLGjAZAIX7tVElKKSWlQVm2pJJi5LvrhNeH8fAbk
# zEG2Qr9CZpTQ/n6obU4t01uitnwWPNh6TcOwwdyJRmPy24SwOCQErwNYJDyi9gLZ
# 7Ni9zHm3nziuLQrzST68aSgrDHulkdjd94xrw2gDiFTT6zA6UVN850w=
# SIG # End signature block
