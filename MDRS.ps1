$filePath = 'C:\Users\jlandry\Documents\Research\Sample - Superstore.xls'
#$sheetName = "Master Document Register"
#$fileName = "xxxx-Master Document Register"
$objExcel = New-Object -ComObject Excel.Application
#$objExcel.Visible = $false
$WorkBook = $objExcel.Workbooks.Open($filePath)

$WorkSheet = $Workbook.Sheets.Item(1)
Write-Host $WorkSheet.Name

$Found = $WorkSheet.Cells.Find('Nebraska')

   
Do {
    $Found = $WorkSheet.Cells.FindNext($Found)
    $Address = $Found.Address(0,0,1,1)
    If ($Address -eq $BeginAddress) {
        BREAK
    }
    [pscustomobject]@{
        WorkSheet = $Worksheet.Name
        Column = $Found.Column
        Row =$Found.Row
        Text = $Found.Text
        Address = $Address
    }                 

   }Until ($false)