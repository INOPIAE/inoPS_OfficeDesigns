Param(
    [Parameter(Mandatory=$True)]
    [string]$xmlFileOutside,

    [Parameter(Mandatory=$True)]
    [string]$originalFile,

    [Parameter(Mandatory=$False)]
    [string]$xmlFileNameInside = 'theme1.xml'
)



$zipFs = New-Object -TypeName System.IO.FileStream -ArgumentList $originalFile, Open
$archive = New-Object System.IO.Compression.ZipArchive -ArgumentList $zipFs, Update
# delete old files
$files = $xmlFileNameInside
($archive.Entries | ? { $files -contains $_.Name }) | % { $_.Delete() }
# add new file
$newEntry = $archive.CreateEntry("word\theme\" + $xmlFileNameInside)
$newFileFs = New-Object -TypeName System.IO.FileStream -ArgumentList $xmlFileOutside, Open
$archiveStream = $newEntry.Open()
$newFileFs.CopyTo($archiveStream)
$archiveStream.Dispose()
# cleanup
$newFileFs.Dispose()
$archive.Dispose()
$zipFs.Dispose()
