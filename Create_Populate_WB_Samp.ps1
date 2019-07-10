$Excel = New-Object -ComObject "Excel.Application"
$Excel.Visible = $true
$workbook = $Excel.Workbooks.Add()
$Sheet = $Workbook.Worksheets.Item("Sheet1")
$Row = 1
$Column = 1
$Sheet.Cells.Item($Row,$Column) = "First Column"
$Column++
$Sheet.Cells.Item($Row,$Column) = "Second Column"
 
$SysDateObject = new-object system.globalization.datetimeformatinfo
$DayNames = $SysDateObject.Daynames
$MonthNames = $SysDateObject.MonthNames
$Column = 1
$Row = 2
$DayNames | %{
   $Sheet.Cells.Item($Row,$Column) = $_
   $Row++
  }
$Column = 2
$Row = 2
$MonthNames | %{
   $Sheet.Cells.Item($Row,$Column) = $_
   $Row++
  }
 
$Range = $Sheet.Range("A1:B1")
$Range.Interior.ColorIndex = 19
$Range.Font.ColorIndex = 11
$Range.Font.Bold = $True   
$Sheet.UsedRange.EntireColumn.AutoFit()
$Excel.ActiveWorkbook.SaveAs("C:\temp\myworkbook.xlsx")