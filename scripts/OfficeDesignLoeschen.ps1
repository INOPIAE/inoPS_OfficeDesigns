
Param(
    [Parameter(Mandatory=$True)]
    [string]$keepDesign
)

# Standardvorlagen entfernen
# $keepDesign = "Unternehmen.*"  bzw $keepDesign = "Unternehmen1.*, Unternehmen2.*"

Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\*.thmx' -Exclude $keepDesign
Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Theme Colors\*.xml' -Exclude $keepDesign
Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Theme Fonts\*.xml'  -Exclude $keepDesign
#Remove-Item -Path 'C:\Program Files (x86)\Microsoft Office\root\Document Themes 16\Theme Effects\*.eftx' 

Remove-Item -Path 'C:\Users\User\AppData\Roaming\Microsoft\Templates\LiveContent\16\Managed\Document Themes\1031\*.thmx'