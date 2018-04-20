
$file = "D:\Suguna\Books.xlsx"
$sheetName = "Sheet1"

$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false

$rowMax = ($sheet.UsedRange.Rows).count

$rowName,$colName = 1,1
$rowParameters,$colParameters = 1,2

for ($i=1; $i -le $rowMax-1; $i++)
{
$name = $sheet.Cells.Item($rowName+$i,$colName).text
$parameters = $sheet.Cells.Item($rowParameters+$i,$colParameters).text

Write-Host ($name+"   "+$parameters)

}

$objExcel.quit()
