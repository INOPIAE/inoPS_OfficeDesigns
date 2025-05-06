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
