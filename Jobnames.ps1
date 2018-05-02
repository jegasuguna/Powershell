#Set-ExecutionPolicy Unrestricted $objWorkbook = $objExcel.Workbooks.Open $Log_path"\ $Log_name"
param([string]$excelfile)
$sheetName = "Sheet1"
Write-Host $PSScriptRoot
$file= Join-Path -Path $excelfile -ChildPath $PSScriptRoot
#Create an instance of Excel.Application and Open Excel file
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item($sheetName)
$objExcel.Visible=$false
#Count max row
$rowMax = ($sheet.UsedRange.Rows).count
#Declare the starting positions
$rowName,$colName = 1,1
$rowParameters,$colParameters = 1,2
#loop to get values and store it
for ($i=1; $i -le $rowMax-1; $i++)
{
$name = $sheet.Cells.Item($rowName+$i,$colName).text
$parameters = $sheet.Cells.Item($rowParameters+$i,$colParameters).text

Write-Host ($name+"   "+$parameters)

}
#close excel file
$objExcel.quit()
