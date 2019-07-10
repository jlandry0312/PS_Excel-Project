$Excel = New-Object -ComObject "Excel.Application"
$Excel.Visible = $false
$workbook = $Excel.Workbooks.Open("C:\Users\jlandry\Documents\Research\xxxx-Master Document Register.xlsm")
$Sheet = $Workbook.Worksheets.Item("Master Document Register")
$Row = 10
$1stColumn = 5
$2ndColumn = 13
Do {
$smDoc = $Sheet.Cells.Item($Row, $1stColumn).Value()
$subDate = $Sheet.Cells.Item($Row, $2ndColumn).Value()
Write-Host "$smDoc,  $subDate"
$Row++
} While ($Sheet.Cells.Item($Row,$2ndColumn).Value() -ne $null)