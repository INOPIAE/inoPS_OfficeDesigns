
Param(
    [Parameter(Mandatory=$True)]
    [string]$keepDesign
)

# Standardvorlagen entfernen
# $keepDesign = "Unternehmen.*"  bzw $keepDesign = "Unternehmen1.*, Unternehmen2.*"

# Office 32 bit

$folderPath = 'C:\Program Files (x86)\Microsoft Office\'

if (Test-Path -Path $folderPath) {
	Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\*.thmx' -Exclude $keepDesign
	Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Theme Colors\*.xml' -Exclude $keepDesign
	Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Theme Fonts\*.xml'  -Exclude $keepDesign
	#Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Theme Effects\*.eftx' 
}

# Office 64 bit

$folderPath = 'C:\Program Files\Microsoft Office\'

if (Test-Path -Path $folderPath) {
	Remove-Item -Path 'C:\Program Files\Microsoft Office\root\Document Themes 16\*.thmx' -Exclude $keepDesign
	Remove-Item -Path 'C:\Program Files\Microsoft Office\root\Document Themes 16\Theme Colors\*.xml' -Exclude $keepDesign
	Remove-Item -Path 'C:\Program Files\Microsoft Office\root\Document Themes 16\Theme Fonts\*.xml'  -Exclude $keepDesign
	#Remove-Item -Path 'C:\Program Files\Microsoft Office\root\Document Themes 16\Theme Effects\*.eftx' 
}

$TemplatesPfad = Join-Path $env:APPDATA 'Microsoft\Templates\LiveContent\16\Managed\Document Themes\1031\*.thmx'
Remove-Item -Path $TemplatesPfad