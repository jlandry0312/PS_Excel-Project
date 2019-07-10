$pathToCsv = "C:\Users\jlandry\Documents\Research\xxxx-Master Document Register_data.csv"
$content = Get-Content -Path $pathToCsv

$content -notmatch'(^[\s,-]*$)|(rows\s*affected)'|Set-Content -Path $pathToCsv
$content | Select-Object -Skip 7 | Set-Content -Path $pathToCsv

$smDWG = @()
$camDWG = @()
$pSubDate = @()

$P = Import-Csv -Path $pathToCsv | ForEach-Object {
    $smDWG += $_."STELLA MARIS"
    $camDWG += $_."CAMERON"
    $pSubDate += $_."PLANNED INI




