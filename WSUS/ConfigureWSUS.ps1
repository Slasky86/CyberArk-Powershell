"*** Configuring WSUS server ***"

"Starting to configure WSUS for Windows updates"

# Variables for the registery path
$WsusRegistryPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate"
$WsusAURegistryPath = "HKLM:\Software\Policies\Microsoft\Windows\WindowsUpdate\AU"
$WsusUpdateRegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"

# Names of the registery variable
$WUServer = "WUServer"
$WUStatusServer = "WUStatusServer"
$AUOptions = "AUOptions"
$RecommendUpdates = "IncludeRecommendedUpdates"
$UseWUServer = "UseWUServer"
$AutoInstallMinorUpdates = "AutoInstallMinorUpdates"
$DetectionFrequencyEnabled = "DetectionFrequencyEnabled"
$DetectionFrequency = "DetectionFrequency"
$NoAutoRebootWithLoggedOnUsers = "NoAutoRebootWithLoggedOnUsers"
$NoAutoUpdate = "NoAutoUpdate"
$ScheduledInstallDay = "ScheduledInstallDay"
$ScheduledInstallTime = "ScheduledInstallTime"

# Names of the services
$wuauservName = "wuauserv"  

Try
{
    # Checking the input
    if ([string]::IsNullOrEmpty($args[0]))
    {
        "Error: WSUS URL parameter is missing"
        return
    }

    #Saving the WSUS server
    $WSUSServer = $args[0]

    Set-Service -Name $wuauservName -StartupType Automatic -Status Running

    IF(!(Test-Path $WsusRegistryPath))
    {
        "Creating the $WsusRegistryPath key"
        New-Item -Path $WsusRegistryPath -Force | Out-Null  
    }

    # Setting the WSUS server to $WSUSServer
    New-ItemProperty -Path $WsusRegistryPath -Name $WUServer -Value $WSUSServer -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $WsusRegistryPath -Name $WUStatusServer -Value $WSUSServer -PropertyType String -Force | Out-Null

    IF(!(Test-Path $WsusAURegistryPath))
    {
        "Creating the $WsusAURegistryPath key"
        New-Item -Path $WsusAURegistryPath -Force | Out-Null
    }

    # Adding the required attributes
    New-ItemProperty -Path $WsusAURegistryPath -Name $AutoInstallMinorUpdates -Value 0 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $DetectionFrequencyEnabled -Value 1 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $DetectionFrequency -Value 14 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $NoAutoRebootWithLoggedOnUsers -Value 1 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $NoAutoUpdate -Value 1 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $AUOptions -Value 5 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $ScheduledInstallDay -Value 0 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $ScheduledInstallTime -Value 14 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusAURegistryPath -Name $UseWUServer -Value 1 -PropertyType DWORD -Force | Out-Null

    IF(!(Test-Path $WsusUpdateRegistryPath))
    {
        "Creating the $WsusUpdateRegistryPath key"
        New-Item -Path $WsusUpdateRegistryPath -Force | Out-Null
    }

    New-ItemProperty -Path $WsusUpdateRegistryPath -Name $AUOptions -Value 1 -PropertyType DWORD -Force | Out-Null
    New-ItemProperty -Path $WsusUpdateRegistryPath -Name $RecommendUpdates -Value 0 -PropertyType DWORD -Force | Out-Null

    #Stoping & Disabling "Windows Update" service
    Get-Service -Name $wuauservName | Stop-Service -Force 
    Set-Service -Name $wuauservName -StartupType Disabled

    "WSUS was configured succesfully"
}
Catch 
{
    "Something is wrong, failed to configure WSUS server."
    return
}

""
"*** Configuring WSUS server Finised ***"
# SIG # Begin signature block
# MIIfdgYJKoZIhvcNAQcCoIIfZzCCH2MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCADExPfz4OHKgMf
# eKP+osXo8yYrBW5KnxYvzIdEg6iHX6CCDnUwggROMIIDNqADAgECAg0B7l8Wnf+X
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
# KDHmv37ZRpQvuxAyctrTAPA6HJtuEJnIo6DhFR9LfTGCEFcwghBTAgEBMH4wbjEL
# MAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExRDBCBgNVBAMT
# O0dsb2JhbFNpZ24gRXh0ZW5kZWQgVmFsaWRhdGlvbiBDb2RlU2lnbmluZyBDQSAt
# IFNIQTI1NiAtIEczAgwhXYQh+9kPSKH6QS4wDQYJYIZIAWUDBAIBBQCgfDAQBgor
# BgEEAYI3AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgBq5E+Ni3i+eI
# LBsyjwJNvWk0wUZZErPiQ9deCVI2aDIwDQYJKoZIhvcNAQEBBQAEggEAifoUA28C
# i6BBVdAtx4aQPDYhO0m3oaBwJFCDBZSE4qPqu85yEea6Fhxd5HFpNkPBeGJjfDFX
# UhpZKLwwBIm2ufaWSN9sNPjd7S8pm3+aw+JyoPcUO/KXa2MKGfYSLMcttkctx7U3
# 72/Chnzumcnmiu0CbAsWWSnJuVTQi8wd5JO26ZXUY3G6scbNgo6cSGyymbgiWP8H
# lxRpFxgPnZgr9Yb9iQ9LcRha1/AkLKyYjNgT8nvYVxS6HC4bLnG7sUr3iMerlmJa
# 9rXwJFuEYs9WhUzvhrubGs0TrwAE2I7GmARpBxM+vdQlKt2gnLuNzjQ4KtsofJzJ
# o57df4ZrbJaOgKGCDiwwgg4oBgorBgEEAYI3AwMBMYIOGDCCDhQGCSqGSIb3DQEH
# AqCCDgUwgg4BAgEDMQ0wCwYJYIZIAWUDBAIBMIH/BgsqhkiG9w0BCRABBKCB7wSB
# 7DCB6QIBAQYLYIZIAYb4RQEHFwMwITAJBgUrDgMCGgUABBQ6WRTuecyLRp6wiBN0
# IrTfMEUiBQIVAIl+NJTfDuOZFZJMPLIKQBf6vrbBGA8yMDIwMDcyMDE3NTQzOFow
# AwIBHqCBhqSBgzCBgDELMAkGA1UEBhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENv
# cnBvcmF0aW9uMR8wHQYDVQQLExZTeW1hbnRlYyBUcnVzdCBOZXR3b3JrMTEwLwYD
# VQQDEyhTeW1hbnRlYyBTSEEyNTYgVGltZVN0YW1waW5nIFNpZ25lciAtIEczoIIK
# izCCBTgwggQgoAMCAQICEHsFsdRJaFFE98mJ0pwZnRIwDQYJKoZIhvcNAQELBQAw
# gb0xCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5WZXJpU2lnbiwgSW5jLjEfMB0GA1UE
# CxMWVmVyaVNpZ24gVHJ1c3QgTmV0d29yazE6MDgGA1UECxMxKGMpIDIwMDggVmVy
# aVNpZ24sIEluYy4gLSBGb3IgYXV0aG9yaXplZCB1c2Ugb25seTE4MDYGA1UEAxMv
# VmVyaVNpZ24gVW5pdmVyc2FsIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkw
# HhcNMTYwMTEyMDAwMDAwWhcNMzEwMTExMjM1OTU5WjB3MQswCQYDVQQGEwJVUzEd
# MBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xHzAdBgNVBAsTFlN5bWFudGVj
# IFRydXN0IE5ldHdvcmsxKDAmBgNVBAMTH1N5bWFudGVjIFNIQTI1NiBUaW1lU3Rh
# bXBpbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC7WZ1ZVU+d
# jHJdGoGi61XzsAGtPHGsMo8Fa4aaJwAyl2pNyWQUSym7wtkpuS7sY7Phzz8LVpD4
# Yht+66YH4t5/Xm1AONSRBudBfHkcy8utG7/YlZHz8O5s+K2WOS5/wSe4eDnFhKXt
# 7a+Hjs6Nx23q0pi1Oh8eOZ3D9Jqo9IThxNF8ccYGKbQ/5IMNJsN7CD5N+Qq3M0n/
# yjvU9bKbS+GImRr1wOkzFNbfx4Dbke7+vJJXcnf0zajM/gn1kze+lYhqxdz0sUvU
# zugJkV+1hHk1inisGTKPI8EyQRtZDqk+scz51ivvt9jk1R1tETqS9pPJnONI7rtT
# DtQ2l4Z4xaE3AgMBAAGjggF3MIIBczAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/
# BAgwBgEB/wIBADBmBgNVHSAEXzBdMFsGC2CGSAGG+EUBBxcDMEwwIwYIKwYBBQUH
# AgEWF2h0dHBzOi8vZC5zeW1jYi5jb20vY3BzMCUGCCsGAQUFBwICMBkaF2h0dHBz
# Oi8vZC5zeW1jYi5jb20vcnBhMC4GCCsGAQUFBwEBBCIwIDAeBggrBgEFBQcwAYYS
# aHR0cDovL3Muc3ltY2QuY29tMDYGA1UdHwQvMC0wK6ApoCeGJWh0dHA6Ly9zLnN5
# bWNiLmNvbS91bml2ZXJzYWwtcm9vdC5jcmwwEwYDVR0lBAwwCgYIKwYBBQUHAwgw
# KAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFtcC0yMDQ4LTMwHQYDVR0O
# BBYEFK9j1sqjToVy4Ke8QfMpojh/gHViMB8GA1UdIwQYMBaAFLZ3+mlIR59TEtXC
# 6gcydgfRlwcZMA0GCSqGSIb3DQEBCwUAA4IBAQB16rAt1TQZXDJF/g7h1E+meMFv
# 1+rd3E/zociBiPenjxXmQCmt5l30otlWZIRxMCrdHmEXZiBWBpgZjV1x8viXvAn9
# HJFHyeLojQP7zJAv1gpsTjPs1rSTyEyQY0g5QCHE3dZuiZg8tZiX6KkGtwnJj1NX
# QZAv4R5NTtzKEHhsQm7wtsX4YVxS9U72a433Snq+8839A9fZ9gOoD+NT9wp17MZ1
# LqpmhQSZt/gGV+HGDvbor9rsmxgfqrnjOgC/zoqUywHbnsc4uw9Sq9HjlANgCk2g
# /idtFDL8P5dA4b+ZidvkORS92uTTw+orWrOVWFUEfcea7CMDjYUq0v+uqWGBMIIF
# SzCCBDOgAwIBAgIQe9Tlr7rMBz+hASMEIkFNEjANBgkqhkiG9w0BAQsFADB3MQsw
# CQYDVQQGEwJVUzEdMBsGA1UEChMUU3ltYW50ZWMgQ29ycG9yYXRpb24xHzAdBgNV
# BAsTFlN5bWFudGVjIFRydXN0IE5ldHdvcmsxKDAmBgNVBAMTH1N5bWFudGVjIFNI
# QTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMTcxMjIzMDAwMDAwWhcNMjkwMzIyMjM1
# OTU5WjCBgDELMAkGA1UEBhMCVVMxHTAbBgNVBAoTFFN5bWFudGVjIENvcnBvcmF0
# aW9uMR8wHQYDVQQLExZTeW1hbnRlYyBUcnVzdCBOZXR3b3JrMTEwLwYDVQQDEyhT
# eW1hbnRlYyBTSEEyNTYgVGltZVN0YW1waW5nIFNpZ25lciAtIEczMIIBIjANBgkq
# hkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEArw6Kqvjcv2l7VBdxRwm9jTyB+HQVd2eQ
# nP3eTgKeS3b25TY+ZdUkIG0w+d0dg+k/J0ozTm0WiuSNQI0iqr6nCxvSB7Y8tRok
# KPgbclE9yAmIJgg6+fpDI3VHcAyzX1uPCB1ySFdlTa8CPED39N0yOJM/5Sym81kj
# y4DeE035EMmqChhsVWFX0fECLMS1q/JsI9KfDQ8ZbK2FYmn9ToXBilIxq1vYyXRS
# 41dsIr9Vf2/KBqs/SrcidmXs7DbylpWBJiz9u5iqATjTryVAmwlT8ClXhVhe6oVI
# QSGH5d600yaye0BTWHmOUjEGTZQDRcTOPAPstwDyOiLFtG/l77CKmwIDAQABo4IB
# xzCCAcMwDAYDVR0TAQH/BAIwADBmBgNVHSAEXzBdMFsGC2CGSAGG+EUBBxcDMEww
# IwYIKwYBBQUHAgEWF2h0dHBzOi8vZC5zeW1jYi5jb20vY3BzMCUGCCsGAQUFBwIC
# MBkaF2h0dHBzOi8vZC5zeW1jYi5jb20vcnBhMEAGA1UdHwQ5MDcwNaAzoDGGL2h0
# dHA6Ly90cy1jcmwud3Muc3ltYW50ZWMuY29tL3NoYTI1Ni10c3MtY2EuY3JsMBYG
# A1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQEAwIHgDB3BggrBgEFBQcB
# AQRrMGkwKgYIKwYBBQUHMAGGHmh0dHA6Ly90cy1vY3NwLndzLnN5bWFudGVjLmNv
# bTA7BggrBgEFBQcwAoYvaHR0cDovL3RzLWFpYS53cy5zeW1hbnRlYy5jb20vc2hh
# MjU2LXRzcy1jYS5jZXIwKAYDVR0RBCEwH6QdMBsxGTAXBgNVBAMTEFRpbWVTdGFt
# cC0yMDQ4LTYwHQYDVR0OBBYEFKUTAamfhcwbbhYeXzsxqnk2AHsdMB8GA1UdIwQY
# MBaAFK9j1sqjToVy4Ke8QfMpojh/gHViMA0GCSqGSIb3DQEBCwUAA4IBAQBGnq/w
# uKJfoplIz6gnSyHNsrmmcnBjL+NVKXs5Rk7nfmUGWIu8V4qSDQjYELo2JPoKe/s7
# 02K/SpQV5oLbilRt/yj+Z89xP+YzCdmiWRD0Hkr+Zcze1GvjUil1AEorpczLm+ip
# Tfe0F1mSQcO3P4bm9sB/RDxGXBda46Q71Wkm1SF94YBnfmKst04uFZrlnCOvWxHq
# calB+Q15OKmhDc+0sdo+mnrHIsV0zd9HCYbE/JElshuW6YUI6N3qdGBuYKVWeg3I
# RFjc5vlIFJ7lv94AvXexmBRyFCTfxxEsHwA/w0sUxmcczB4Go5BfXFSLPuMzW4IP
# xbeGAk5xn+lmRT92MYICWjCCAlYCAQEwgYswdzELMAkGA1UEBhMCVVMxHTAbBgNV
# BAoTFFN5bWFudGVjIENvcnBvcmF0aW9uMR8wHQYDVQQLExZTeW1hbnRlYyBUcnVz
# dCBOZXR3b3JrMSgwJgYDVQQDEx9TeW1hbnRlYyBTSEEyNTYgVGltZVN0YW1waW5n
# IENBAhB71OWvuswHP6EBIwQiQU0SMAsGCWCGSAFlAwQCAaCBpDAaBgkqhkiG9w0B
# CQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTIwMDcyMDE3NTQzOFow
# LwYJKoZIhvcNAQkEMSIEIHCBBs4+cDL1Pp1+jZ484L2vHynNyGasKGIOkJtWz9gX
# MDcGCyqGSIb3DQEJEAIvMSgwJjAkMCIEIMR0znYAfQI5Tg2l5N58FMaA+eKCATz+
# 9lPvXbcf32H4MAsGCSqGSIb3DQEBAQSCAQCA0H0mSpecCl+cVm8tpsfysceALwBh
# /jsIzSgzONw/hczSTcVsupM/p7+EKIyt4LUfEIen10423XDBVQpANwZPDXSMRjaD
# bgL6q0iBoUvjC1PaL8eDIFoGWo2wYPmaEKyPUOKHQ0o03EVwnqFSvrs7XJ2sLukT
# B1ugrdU/TFQCgM+rc7SC6QAoWOQtxrmXxRZffM03WuRr7yYkOjsCktQ2uUkJQnD8
# 1+f3vcij9dGx4of1eE1ppVaNvDN4RSucs1ZRmBDke2h5Y6+A6zkJjlSUouNFuepN
# Z8kn0If1zHF+Krb+WGjd+VUsXWrpxvtHiPz1/Hpp4goDQBDUclDrX/MA
# SIG # End signature block
