Write-Host ""

# Create folder to store output files if it doesn't exist
$storefolder = "$HOME\.cranfield-diags"
if(!(Test-Path -Path $storefolder )){
    New-Item -ItemType directory -Path $storefolder
}

# Define the variables we'll use going forward
$dt = Get-Date -Format FileDateTimeUniversal | ForEach-Object { $_ -replace ":", "." }
$outputfilename = "$dt-$env:COMPUTERNAME-netdiags.txt"
$outputfilepath = "$storefolder\$outputfilename"
$seperator = "----------------------------------------------------------------------------------------------------"
$webhost = "www.google.com"
$googledns1 = "8.8.8.8"
$cloudflaredns1 = "1.1.1.1"
$akamaidns1 = "88.221.162.98"
$akamaidns2 = "88.221.163.98"
$mscontesturl = "http://www.msftncsi.com/ncsi.txt"
$ip = Invoke-RestMethod -uri https://icanhazip.com/s

# Define tests
$hash = $null
$hash = [ordered]@{}
$hash.add("Retrieving host and user details",'Write-Output "Hostname: $env:COMPUTERNAME`nUsername: $env:USERNAME`nUserdomain: $env:USERDOMAIN"')
$hash.add("Retrieving last boot time",'Get-CimInstance Win32_OperatingSystem | Select-Object LastBootUpTime')
$hash.add("Retrieving Windows version",'[System.Environment]::OSVersion.Version')
$hash.add("Get personal certificates",'Get-ChildItem -path Cert:\CurrentUser\My | Select-Object *')
$hash.add("Retrieving Windows activation status",'(cscript /Nologo "C:\Windows\System32\slmgr.vbs" /xpr) -join ''''')
$hash.add("Retrieving disk space",'Get-CimInstance -Class CIM_LogicalDisk | Select-Object @{Name="Size(GB)";Expression={$_.size/1gb}}, @{Name="Free Space(GB)";Expression={$_.freespace/1gb}}, @{Name="Free (%)";Expression={"{0,6:P0}" -f(($_.freespace/1gb) / ($_.size/1gb))}}, DeviceID, DriveType | Where-Object DriveType -EQ ''3''')
$hash.add("Retrieving IP configuration",'Get-NetIPConfiguration -All')
$hash.add("Running basic network connectivity test",'Test-NetConnection -ComputerName "$webhost" -InformationLevel "Detailed"')
$hash.add("Running ping tests",'Test-Connection -ComputerName $googledns1,$cloudflaredns1,$akamaidns1,$akamaidns2 -Count 1')
$hash.add("Running resolving test: $googledns1",'Resolve-DnsName microsoft.com -Server $googledns1 -QuickTimeout')
$hash.add("Running resolving test: $cloudflaredns1",'Resolve-DnsName microsoft.com -Server $cloudflaredns1 -QuickTimeout')
$hash.add("Running resolving test: $akamaidns1",'Resolve-DnsName microsoft.com -Server $akamaidns1 -QuickTimeout')
$hash.add("Running resolving test: $akamaidns2",'Resolve-DnsName microsoft.com -Server $akamaidns2 -QuickTimeout')
$hash.add("Running ipconfig/all",'ipconfig /all')
$hash.add("Retrieving current wireless network",'netsh wlan show interfaces')
$hash.add("Retrieving all visible wireless networks",'Get-WifiNetwork')
$hash.add("Retrieving all saved wireless networks",'netsh wlan show profiles')
$hash.add("Retrieving Resnet Personal wireless details (if saved)",'netsh wlan show profiles name="Resnet Personal"')
$hash.add("Retrieving Eduroam wireless details (if saved)",'netsh wlan show profiles name="Eduroam"')
$hash.add("Retrieving the MS connection test page",'Invoke-WebRequest $mscontesturl -UseBasicParsing')
$hash.add("Retrieving service states",'Get-Service')
$hash.add("Retrieving join state",'dsregcmd /status')

# Create functions we'll use later
Function outtofile ($P1)
{
    $p1 | Out-File -FilePath $outputfilepath -Append
    Write-Host "....."
}

Function outtofilespacer
{
    "" | Out-File -FilePath $outputfilepath -Append
    "**************************************************************************************************************" | Out-File -FilePath $outputfilepath -Append
    "" | Out-File -FilePath $outputfilepath -Append
}

Function outtofiledebug ($P1)
{
    ">> $p1" | Out-File -FilePath $outputfilepath -Append
    Write-Host ">> $p1"
}

Function outtofilecode ($P1)
{
    "   **$p1**" | Out-File -FilePath $outputfilepath -Append
    "" | Out-File -FilePath $outputfilepath -Append
}

Function Get-WifiNetwork
{
    end {
        netsh wlan sh net mode=bssid | % -process {
            if ($_ -match '^SSID (\d+) : (.*)$') {
                $current = @{}
                $networks += $current
                $current.Index = $matches[1].trim()
                $current.SSID = $matches[2].trim()
            } else {
                if ($_ -match '^\s+(.*)\s+:\s+(.*)\s*$') {
                    $current[$matches[1].trim()] = $matches[2].trim()
                }
            }
        } -begin { $networks = @() } -end { $networks|% { new-object psobject -property $_ } }
    }
}

# Start the diagnostics
Write-Host $seperator
Write-Host ""
outtofiledebug("Diagnostics started: $dt - writing to $outputfilepath")
Write-Host ""
foreach ($h in $hash.GetEnumerator()) {
    outtofilespacer
    outtofiledebug "$($h.Name)"
    outtofilecode "$($h.Value)"
    $testresult = Invoke-Expression $($h.Value)
    outtofile($testresult)
    }

Write-Host ">> Tests completed"
Write-Host ""
Write-Host "$seperator"
Write-Host "Results location:"
Write-Host "  Disk: $outputfilepath"
Write-Host "$seperator"
Write-Host ""
Invoke-Item $storefolder
# SIG # Begin signature block
# MIIPZAYJKoZIhvcNAQcCoIIPVTCCD1ECAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUE0biHOBV64sSMcvWU3UJ5dxz
# PcegggySMIIFkzCCBHugAwIBAgITawAAAAIwIVocx1eJngAAAAAAAjANBgkqhkiG
# 9w0BAQsFADAnMSUwIwYDVQQDExxDcmFuZmllbGQgVW5pdmVyc2l0eSBSb290IENB
# MB4XDTE4MTAxOTA4MjIzNVoXDTMzMTAxOTA4MzIzNVowgYIxEjAQBgoJkiaJk/Is
# ZAEZFgJ1azESMBAGCgmSJomT8ixkARkWAmFjMRkwFwYKCZImiZPyLGQBGRYJY3Jh
# bmZpZWxkMRMwEQYKCZImiZPyLGQBGRYDY25zMSgwJgYDVQQDEx9DcmFuZmllbGQg
# VW5pdmVyc2l0eSBJc3N1aW5nIENBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEA3enErZqp++OB9+1f9UU10q5Xl/x8PTS5x5xuUX0oCwKMjQnH2/az5oYH
# /u5UaqKW4kcbeZy01SlrtLks20+zdBWuJCVXNTqw9bms4/l4hVQNsd+YiXDbeeA3
# 4tUDvCnifcpzRFvBXRU4Nc5FUsMYwYAjfmua4tEylgjOXdFvlGUjd8TSGZS6iH08
# Jn2Qubq8vLfEc1cFYTnBezkHYUd8J/SsbsatvUbkv1ZF8XWYdorlQyIKaDGXfpME
# lY7p3lLmwZnVZhapMPu3D96+hAthoMS0SfJTDUd5pnsnwngl+rnxNVe2Z9bGjVkU
# T0ggaNxivRfXXUixR2YttcYuuauxkhyp13iVU8Zc4Vttlh3Pv5Nv6IejehsfCZlH
# xOkdktREPBDara7EyFZ4RMQ53PJSoGhAxSQ+QhEv9zShsX4Uwtr+ougdA/etuT8X
# 8V0XeR7G5JLcG9vqw1LuKk5+trW7y1o19ZaN1rZx9k09ysVemDopiXPeexf1U7Mp
# MWoGirtEmMjc/u9ZfgRjAEdv1/4zQGqX5754etIl1vqMwi3SgtdajqT/8y0Th5CW
# tixvMepf52qhCNHVMMtqtP/CxsWQrF2CVgEw/QIcFL8npRGoBj7QIH4uRtealXYW
# lMgVHp1swYG1UrsZ3wXDe552hTUUZI/LHneh2THr/Olwn+3cbN0CAwEAAaOCAVow
# ggFWMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBSmXsrnNKOMeYiraQDLW1fR
# /rrXFzAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBQ+aX8qeXfBMXB1RPwnV98y7mrfoTBN
# BgNVHR8ERjBEMEKgQKA+hjxodHRwOi8vcGtpLmNyYW5maWVsZC5hYy51ay9wa2kv
# Q1JBTkZJRUxEX1VOSVZFUlNJVFktUk9PVC5jcmwwegYIKwYBBQUHAQEEbjBsMGoG
# CCsGAQUFBzAChl5odHRwOi8vcGtpLmNyYW5maWVsZC5hYy51ay9wa2kvQ1JBTkZJ
# RUxEX1VOSVZFUlNJVFktUk9PVENyYW5maWVsZCUyMFVuaXZlcnNpdHklMjBSb290
# JTIwQ0EuY3J0MA0GCSqGSIb3DQEBCwUAA4IBAQCzJZCeSB5x5J3McJ2AVMAa9iBg
# vdiKSr6LLqCY6avlXdyAwMfsrRHHvcM8g677BfVBcFSLi8KI7kdGMOXTtNOjoe13
# eIQY8y9I3CaqB0TLsFCiB26hG+nrIK3BgunhJxMz3OksSbYOaR0lS4VS1i9tWnEC
# tixJm1i7U4cugtU2/DZ6KM5osWCgD8zAhfQEAZV5f6W4n7HXQC8COvL+sfv1ti1j
# 4we9/KXU9jLfwYTgdqTNVklH3+ME4pUzCgQMnTnMx6U/VrVtvvjs8wYTq4mBW7vK
# ZDy+H8v1XN2kNS4l5uVBrTBX2je9YgEs9io29+Dv4b/jCzilhiLjpcJtYqTGMIIG
# 9zCCBN+gAwIBAgITZgABEQcnYOOZh7btPAAAAAERBzANBgkqhkiG9w0BAQsFADCB
# gjESMBAGCgmSJomT8ixkARkWAnVrMRIwEAYKCZImiZPyLGQBGRYCYWMxGTAXBgoJ
# kiaJk/IsZAEZFgljcmFuZmllbGQxEzARBgoJkiaJk/IsZAEZFgNjbnMxKDAmBgNV
# BAMTH0NyYW5maWVsZCBVbml2ZXJzaXR5IElzc3VpbmcgQ0EwHhcNMjAxMDAyMTUw
# ODE2WhcNMjUxMDAxMTUwODE2WjCBpjELMAkGA1UEBhMCR0IxFTATBgNVBAgTDEJl
# ZGZvcmRzaGlyZTEQMA4GA1UEBxMHQmVkZm9yZDEdMBsGA1UEChMUQ3JhbmZpZWxk
# IFVuaXZlcnNpdHkxJjAkBgNVBAsTHUlUIC0gQ29kZSBTaWduaW5nIENlcnRpZmlj
# YXRlMScwJQYDVQQDDB5sdWtlLndoaXR3b3J0aEBjcmFuZmllbGQuYWMudWswggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCeBdORLkIyjXMe54naF04UXFdE
# L5Ot1RP0EHyuJvft2Rp/BTneG+vbCem1ZU3I3MipKjtLXJXMggjNqbb/TNSN2+Z/
# 0KUiR40txlTaAttHD2ZGqs0sIDWKlUPLtjpkooQKBVHRPtS24+HDj1mizKx/fsvq
# UDZGTLoryAlo1wIgwNxjXQECpAJKliskKiR2SlyrGTkszNPIf8IbY386jSN42jfo
# IOHCuuCpw7RwsddiA28v4+pUKdBn0hynuJ2QYZxiiWBTSsB/Tzr1YlGohcZHUkPj
# XXl5GahvMjhNpDvKotJr4rhea2JorzxqkTIJIU/++p4mWsbToUmQpS4v8cDrAgMB
# AAGjggI+MIICOjA9BgkrBgEEAYI3FQcEMDAuBiYrBgEEAYI3FQiCvs4HgfnIeuWH
# HYTD6GqC/cpMgSKFr+U1hcLVUwIBZAIBAzATBgNVHSUEDDAKBggrBgEFBQcDAzAL
# BgNVHQ8EBAMCB4AwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzBEBgkqhkiG
# 9w0BCQ8ENzA1MA4GCCqGSIb3DQMCAgIAgDAOBggqhkiG9w0DBAICAIAwBwYFKw4D
# AgcwCgYIKoZIhvcNAwcwHQYDVR0OBBYEFDlwI637ZZBVMVB3ITwwOOVYdG1HMCkG
# A1UdEQQiMCCBHmx1a2Uud2hpdHdvcnRoQGNyYW5maWVsZC5hYy51azAfBgNVHSME
# GDAWgBSmXsrnNKOMeYiraQDLW1fR/rrXFzBMBgNVHR8ERTBDMEGgP6A9hjtodHRw
# Oi8vcGtpLmNyYW5maWVsZC5hYy51ay9wa2kvQ1JBTkZJRUxEX1VOSVZFUlNJVFkt
# RUNBLmNybDCBugYIKwYBBQUHAQEEga0wgaowbAYIKwYBBQUHMAKGYGh0dHA6Ly9w
# a2kuY3JhbmZpZWxkLmFjLnVrL3BraS9DUkFORklFTERfVU5JVkVSU0lUWS1FQ0FD
# cmFuZmllbGQlMjBVbml2ZXJzaXR5JTIwSXNzdWluZyUyMENBLmNydDA6BggrBgEF
# BQcwAYYuaHR0cDovL2NjcGtpd2ViLTEuY2VudHJhbC5jcmFuZmllbGQuYWMudWsv
# b2NzcDANBgkqhkiG9w0BAQsFAAOCAgEAtTHSoZ2NwTn3yyC7DsSS7UHRLCzZe7ng
# m0n1ByySmXM2hElgg2JdAfA0Qm6iFICTRyYz5VGXzsQC8cHDlk1bzMiS4YDklsEw
# cM9w4B8cg91qOVAoLElIBXvr6W0mvCIBOWxEfAJQZx0IlLQyX8sAfGM6S77qrzEC
# It9mM/Ff6UVS8vwEbgABQv6zwOWSRPFPr0QTzxgQ+cMdxBs4/m/JIwEPvdzaqphG
# 6AYl8MCBy0kRfzi3/syugAUePRwQ/s6zwtejSg8tCcUGw94TuSm1qHewAkJ1/IFe
# Cdw1Ntjs4hI/d9GDiMYL30gGKTbK0/gFs3DKnlpXusu5RwtKXncTsm2rz3o1LO0x
# LB3FvkKVweYVFi2TPOitlXQXIfp8thUygknn18gxqT3t0/Q2jDaaYaYLJV443mUC
# SnMAFOBgHyhm6ZdJaB+vTHdpkxb+OfZLFr8/aqnU/i1V5nO03auAIZyrVKSeBf1X
# kcBAaHTijDZrPF5FHm8wYvkL+hZO39hHUo2x+BlBAEplAJbfJ16PKuTA97mX8ACi
# lue8g9vAh2+McIFxu6DQ8toUXCyjeRpC9abSAYJG76fFKOForvEQh2MdcmYFQj+H
# 4UxIYeBvJwItVn5fXbfXfQnr+LaP1ZlBVYVpKerR2npnqor7QEnCWyrufVFVOZoV
# 1qVta7H3Z78xggI8MIICOAIBATCBmjCBgjESMBAGCgmSJomT8ixkARkWAnVrMRIw
# EAYKCZImiZPyLGQBGRYCYWMxGTAXBgoJkiaJk/IsZAEZFgljcmFuZmllbGQxEzAR
# BgoJkiaJk/IsZAEZFgNjbnMxKDAmBgNVBAMTH0NyYW5maWVsZCBVbml2ZXJzaXR5
# IElzc3VpbmcgQ0ECE2YAAREHJ2DjmYe27TwAAAABEQcwCQYFKw4DAhoFAKB4MBgG
# CisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcC
# AQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYE
# FBc5/0JwrCJxwQHTEjndaU+D24SsMA0GCSqGSIb3DQEBAQUABIIBADZJAoooUayN
# 6jL29N5tv0pG0UU2hCZvn4cMuM8tL2iBGqIOTrUUP4tAJ3Q7fkX9tOQdLyLwjXlv
# ltBeJBn/YqgsPEgjqBQcnv95O/USVH+sKMGlsJOmaR/u6CVBMD7CB7o/urKTXmTe
# gkvGfb+bdrRgNkX3KJelqloxwxHFDrLx+j//5gEWeH1X+5BHiJ4QkT+U468nmf5T
# HHMsWtffxoPJxlLYlclMH6zrQTdpLWq8gLC8AFNtHZ5LT71TYId6ZUqILEIE1f9z
# wK2BVXX0nCEdYhwoKprQoNWcArN4c8TWvOTBu+7N80APnfF6k9zIRRzUhzyCPTra
# C/HxoNcAghI=
# SIG # End signature block
