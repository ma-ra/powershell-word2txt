<#
.SYNOPSIS
    Find doc, docx, rtf, odt files and convert them to txt.
.DESCRIPTION
    Find doc, docx, rtf, odt files and convert them to txt. To work needs installed MS Word - tested with MS Word 2016
#>

# config
$FILES_TO_SEARCH="*.doc","*.docx","*.rtf","*.odt"
$WHERE_TO_SEARCH="C:\Users\"
$IGNORE_PATH="*\Dysk Google\*"
$WHERE_SAVE_TXT="$PSScriptRoot\converted_files\"

# find files with exception
Write-Output "##### Wyszukiwanie plików"
Get-ChildItem -Path $WHERE_TO_SEARCH -Recurse -Include $FILES_TO_SEARCH | Where-Object {$_.FullName -NotLike $IGNORE_PATH} | Tee-Object -Variable files | %{$_.FullName}
$count = ($files | measure).Count

# create destination
If(!(test-path $WHERE_SAVE_TXT)) { New-Item -ItemType Directory -Force -Path $WHERE_SAVE_TXT | Out-Null}

# convert files to txt
Write-Output "##### Konwersja plików"
$wordApplication = New-Object -ComObject Word.Application
$wdSaveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat],    "wdFormatText")
$converted=0

foreach ($file in $files) {
   $converted++ ; Write-Output "[$converted/$count] $($file.Name)"
   $wordDocument = $wordApplication.Documents.Open($file.FullName)
   $wordDocument.SaveAs($WHERE_SAVE_TXT + $file.Name + ".txt", [ref]$wdSaveFormat)
   $WordDocument.Close()
}
$wordApplication.Quit()
