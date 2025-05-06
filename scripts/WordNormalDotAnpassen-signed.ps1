Param(
    [Parameter(Mandatory = $true)]
    [string]$xmlFileOutside,

    [Parameter(Mandatory = $true)]
    [string]$originalFile,

    [Parameter(Mandatory = $false)]
    [string]$xmlFileNameInside = 'theme1.xml'
)

# Validierung
if (!(Test-Path $xmlFileOutside)) {
    Write-Error "Datei $xmlFileOutside nicht gefunden."
    exit 1
}
if (!(Test-Path $originalFile)) {
    Write-Error "Datei $originalFile nicht gefunden."
    exit 1
}

# .NET ZIP-Funktionen bereitstellen
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ZIP öffnen
try {
    $archive = [System.IO.Compression.ZipFile]::Open($originalFile, 'Update')
} catch {
    Write-Error "Konnte Datei nicht als ZIP öffnen: $originalFile"
    exit 1
}

# Zielpfad im Archiv
$zipPath = "word/theme/$xmlFileNameInside"

# ggf. vorhandene Datei im ZIP löschen
$entry = $archive.Entries | Where-Object { $_.FullName -ieq $zipPath }
if ($entry) {
    $entry.Delete()
}

# Neue Datei hinzufügen
try {
    $newEntry = $archive.CreateEntry($zipPath)
    $entryStream = $newEntry.Open()
    $inputStream = [System.IO.File]::OpenRead($xmlFileOutside)
    $inputStream.CopyTo($entryStream)

    # Aufräumen
    $entryStream.Dispose()
    $inputStream.Dispose()
    $archive.Dispose()

    Write-Host "Datei erfolgreich eingefügt: $zipPath"
} catch {
    Write-Error "Fehler beim Einfügen der Datei: $_"
    $archive.Dispose()
    exit 1
}

# SIG # Begin signature block
# MIIOfAYJKoZIhvcNAQcCoIIObTCCDmkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD5pDWOlG+2YPvg
# RnSVFbUXl7cIIXFe8NnrXJ4xfTXvHqCCC8UwggWAMIIDaKADAgECAgMDDbkwDQYJ
# KoZIhvcNAQENBQAwVDEUMBIGA1UEChMLQ0FjZXJ0IEluYy4xHjAcBgNVBAsTFWh0
# dHA6Ly93d3cuQ0FjZXJ0Lm9yZzEcMBoGA1UEAxMTQ0FjZXJ0IENsYXNzIDMgUm9v
# dDAeFw0yNTAxMDMwNjI4MzhaFw0yNzAxMDMwNjI4MzhaMEAxGTAXBgNVBAMMEE1h
# cmN1cyBNw4PCpG5nZWwxIzAhBgkqhkiG9w0BCQEWFG0ubWFlbmdlbEBpbm9waWFl
# LmRlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzxEV/p4FaE6ylTGg
# rlmxljKD3KdZBhe+q10ZWE5HLkPycbmJbDZxIVi5H7MWynAvhmAutV30XVQtpcxr
# rRH20hmODUI8VVwlxvpUeb5GK/n9DSjaD9u9Unt51g8cAT/Ss0vfB38DL7XFxVC5
# ARYKeR45eBLqDhagLaCGvuA3PyOmrBy53KByB+ErONNQQktGCE8nFWJuMqOKjNM7
# KWpShSK6C9NBMjcNnIacOBIMWoZnAMSys6ovJbnv0a+VzTGjtZn/2CbxAhl4zFm7
# bzUBCH711URuw3s3HslYEMB9K24tZzMGVW5rnaQGGN2JUSzpx3BVXT+qSfsgG3RX
# BFaWjQIDAQABo4IBbTCCAWkwDAYDVR0TAQH/BAIwADBWBglghkgBhvhCAQ0ESRZH
# VG8gZ2V0IHlvdXIgb3duIGNlcnRpZmljYXRlIGZvciBGUkVFIGhlYWQgb3ZlciB0
# byBodHRwOi8vd3d3LkNBY2VydC5vcmcwDgYDVR0PAQH/BAQDAgOoMGIGA1UdJQRb
# MFkGCCsGAQUFBwMEBggrBgEFBQcDAgYIKwYBBQUHAwMGCisGAQQBgjcCARUGCisG
# AQQBgjcCARYGCisGAQQBgjcKAwQGCisGAQQBgjcKAwMGCWCGSAGG+EIEATAyBggr
# BgEFBQcBAQQmMCQwIgYIKwYBBQUHMAGGFmh0dHA6Ly9vY3NwLmNhY2VydC5vcmcw
# OAYDVR0fBDEwLzAtoCugKYYnaHR0cDovL2NybC5jYWNlcnQub3JnL2NsYXNzMy1y
# ZXZva2UuY3JsMB8GA1UdEQQYMBaBFG0ubWFlbmdlbEBpbm9waWFlLmRlMA0GCSqG
# SIb3DQEBDQUAA4ICAQCcF71aDoz5CDPS5CT5PPu55WCZdCacJb7ywghtYLyHYr41
# 5hvOEsYQFkjU39vX2oa2B4KNIn3RF9pLDc7UJgqIANH06/NrNKTQM6DD+tVM9gOe
# nprX8aiUXzQ8VbzMS4TUxDBAOFaveFulEvLViT9duQ3ItaSJEbIZrbZAWu5neT83
# xJxvIhaee1H/Y5znnE7TRL6gZ07Yp0pW1uvXF0NpQ3VKXBtV7DtKuPGGtyVUA7oZ
# bSfLdKersfpUspwb5ssoDnEijH6R1WMQZV2JiHv/77BbNE+/3vMSQh3FDRovTcVl
# 7B6JLTa2KCoD5Ls96pmGzbJB3lml8I3Q0BsViL58l2M9CbgaH5AQDTS7exqFCdAm
# O/sLwaxb6foUpzBhgamE2TM8lACZU3zQ1CrEmBz6CWfdVVNdlDTLVdS7YP/YFPFb
# 2+OqQ4AW1xDfki7ChgtpvJ0ojDvQWqs4eWapIkUefxzdtcn+y9q9k5SaU4Cpfvh4
# OtxvQy/3u2D+tAyiDhSISmFy6IxYMNIX7aR2AfrLeo6W8e9HBuYeh7vtJlcJpxaY
# /XQlWp0I0BYR+USb+2Od5wDFs6ZTTBzQCNNn+J02bZccmSULJEHJSZgdKDabSAc6
# 03mB8uqT5Q4j/tV+Pi0UXjBeuLY+ZyLLaL1epAMRNQGP+g+JGJJelGJ8dntyuTCC
# Bj0wggQloAMCAQICAxTiKDANBgkqhkiG9w0BAQ0FADB5MRAwDgYDVQQKEwdSb290
# IENBMR4wHAYDVQQLExVodHRwOi8vd3d3LmNhY2VydC5vcmcxIjAgBgNVBAMTGUNB
# IENlcnQgU2lnbmluZyBBdXRob3JpdHkxITAfBgkqhkiG9w0BCQEWEnN1cHBvcnRA
# Y2FjZXJ0Lm9yZzAeFw0yMTA0MTkxMjE4MzBaFw0zMTA0MTcxMjE4MzBaMFQxFDAS
# BgNVBAoTC0NBY2VydCBJbmMuMR4wHAYDVQQLExVodHRwOi8vd3d3LkNBY2VydC5v
# cmcxHDAaBgNVBAMTE0NBY2VydCBDbGFzcyAzIFJvb3QwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCrSTURSHzSJn5TlM9Dqd0o10Iqi/OHeBlYfA+e2ol9
# 4fvrcpANdKGWZKufoCSZc9riVXbHF3v1BKxGuMO+f2SNEGwk82GcwPKQ+lHm9WkB
# Y8MPVuJKQs/iRIwlKKjFeQl9RrmK8+nzNCkIReQcn8uUBByBqBSzmGXEQ+xOgo0J
# 0b2qW42S0OzekMV/CsLj6+YxWl50PpczWejDAz1gM7/30W9HxM3uYoNSbi4ImqTZ
# FRiRpoWSR7CuSOtttyHshRpocjWr//AQXcD0lKdq1TuSfkyQBX6TwSyLpI5idBVx
# bgtxA+qvFTia1NIFcm+M+SvrWnIl+TlG43IbPgTDZCciECqKT1inA62+tC4T7V2q
# SNfVfdQqe1z6RgRQ5MwOQluM7dvyz/yWk+DbETZUYjQ4jwxgmzuXVjit89Jbi6Bb
# 6k6WuHzX1aCGcEDTkSm3ojyt9Yy7zxqSiuQ0e8DYbF/pCsLDpyCaWt8sXVJcukfV
# m+8kKHA4IC/VfynAskEDaJLM4JzMl0tF7zoQCqtwOpiVcK01seqFK6QcgCExqa5g
# eoAmSAC4AcCTY1UikTxW56/bOiXzjzFU6iaLgVn5odFTEcV7nQP2dBHgbbEsPyyG
# kZlxmqZ3izRg0RS0LKydr4wQ05/EavhvE/xzWfdmQnQeiuP43NJvmJzLR5iVQAX7
# 6QIDAQABo4HyMIHvMA8GA1UdEwEB/wQFMAMBAf8wYQYIKwYBBQUHAQEEVTBTMCMG
# CCsGAQUFBzABhhdodHRwOi8vb2NzcC5DQWNlcnQub3JnLzAsBggrBgEFBQcwAoYg
# aHR0cDovL3d3dy5DQWNlcnQub3JnL2NsYXNzMy5jcnQwRQYDVR0gBD4wPDA6Bgsr
# BgEEAYGQSgIDATArMCkGCCsGAQUFBwIBFh1odHRwOi8vd3d3LkNBY2VydC5vcmcv
# Y3BzLnBocDAyBgNVHR8EKzApMCegJaAjhiFodHRwczovL3d3dy5jYWNlcnQub3Jn
# L2NsYXNzMy5jcmwwDQYJKoZIhvcNAQENBQADggIBAMYerXdctCib0ciNRBLAvXZ2
# BIMhB/gRgn9rwZVCwDgRtSVwjYsMwdVs/RwaA3yL+AYxpZ3eQSnUi5uE1z3BN4Zx
# ox9bYSkeXXd9u/CtuRUZExTmNYD/phm0N4WUQeiIw1/gsgaku/hAqR05rO3qP5gE
# TfmM+Ud5c1L17N80l/s+d+Dc0YOIuohzR1qmpBXEDXAND55LEwd+7xg++aUBqnkp
# sedS+lM6yKZ/tu+JobGhTS/OY4V/pSrpO9TBo88KE4W7mdecZpCE52bUULOh4S0i
# KiVowyCyK8S6mB3oSu9cWMK0TYRW9067FmhCbJK4b3jNDrP776CzZIfy94hEOfy5
# 5izAmCTUQCxeyO4LHbgCTSa4ChjGLx5LdW6PLiFzvMIDVe6qFOCaGgdTC99EFKhn
# Ba9EyNOhRXYCtn8MuYbpT8ZusLsVtL/ogLV2Mf5kZMEKWG3FULayA78dQk9ZOdHE
# MYvoyCo5HBVh8N5AaA5wqLNP7pHoD0+2kJ5NgGy+HO5wpLgHBCsNQQJUhE5H6ouW
# 7XZYYefDIXsGb9S3C+c0MoPMNabnJU98F0L8vFcDxp9Cf5hg+ICy2faxnBw1BAqJ
# MRaFpPruTAnFaphm7MhuKubLktwjbJbB1EXzPG0CuKC7x0fCwhxATEXHRQZ/O3Er
# ziungdZHRiiwPMpl8WafMYICDTCCAgkCAQEwWzBUMRQwEgYDVQQKEwtDQWNlcnQg
# SW5jLjEeMBwGA1UECxMVaHR0cDovL3d3dy5DQWNlcnQub3JnMRwwGgYDVQQDExND
# QWNlcnQgQ2xhc3MgMyBSb290AgMDDbkwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYB
# BAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAc
# BgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFjAvBgkqhkiG9w0BCQQxIgQgcKs4
# ZceHJuZgt9WlQtHLBA2Mx+ot816GJLSJYqgxbDUwDQYJKoZIhvcNAQEBBQAEggEA
# W6Y6/6AwBPU0G/9DlkB4s8V1IMEUi+UMg8hOnv0XZtPpwc/dq51blpzZ/BicwLzX
# QVhC/us3ewSypkb2/TFvQt+Qt3s6/pjTOda0gP8ma7rt1lr7998dWH9QB92Z+nnr
# BF8c8fPMXZB+HFJ4T8RyMAFRwsvS9ojeosrM5NmeVJcCcsuLHKGVm0TLGA3T6x2B
# W4/W0ZrvgoElneW9JouyCviO7b/97tpkNUmKT2TID3mpQZMwaAt+6nY97HKt4A5p
# oIWEEjshs9/yvFb/HwmK7uijHT6DcZybMZ0jzlFn72OzgHs4bbcjMmmEuDiSuFCe
# ihgImd4c4CJc1EmKB1YLTA==
# SIG # End signature block
