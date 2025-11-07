$ErrorActionPreference = 'Stop'
$excelPath = 'E:\SAPTARSHI BANERJEE\WORKLOGS\summary.xlsx'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$worksheet = $workbook.Worksheets.Item(1)
$usedRange = $worksheet.UsedRange
$values = $usedRange.Value2
$workbook.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
$values | ConvertTo-Json