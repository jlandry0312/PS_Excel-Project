$pathToCsv = "C:\Users\Jland\Google Drive\Stella Maris\PS_Excel-Project\PS_Excel-Project\xxxx-Master Document Register_data.csv"
#$content = Get-Content -Path $pathToCsv

#$content -notmatch'(^[\s,-]*$)|(rows\s*affected)'|Set-Content -Path $pathToCsv
#$content | Select-Object -Skip 7 | Set-Content -Path $pathToCsv

$smDWG = @()
$camDWG = @()
$pSubDate = @()

Import-Csv -Path $pathToCsv
ForEach-Object {
    $smDWG += $_."STELLA MARIS DOC. / DRAWING NO."
    #$camDWG += $_."CAMERON"
    #$pSubDate += $_."PLANNED INITIAL SUBMITTAL DATE"
}

#Write-Host $smDWG


