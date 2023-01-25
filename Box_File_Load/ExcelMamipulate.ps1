$Date = (Get-Date).AddDays(-6).ToString('M.dd.yy')
$Sourcefile = "HFS Americas - Deal by Deal Covid Restructuring Tracker - ($Date).xlsx"
Import-Excel $PSScriptRoot"\"$Sourcefile -WorkSheetname 'Americas Restructure by Schedul' -NoHeader | Export-Excel "$PSScriptRoot\CovidData.xlsx" -WorkSheetname Sheet1

$Excel = New-Object -ComObject Excel.Application
$Workbook=$Excel.Workbooks.Open("$PSScriptRoot\CovidData.xlsx")
$WorkSheet = $Workbook.Sheets.Item(1)
$WorkSheet.Columns.Replace("'","''")

#Delete Blank Columns
$WorkSheet.Range("AW1:BK1").EntireColumn.Delete()
$WorkSheet.Range("AO1:AP1").EntireColumn.Delete()
$WorkSheet.Range("AK1:AL1").EntireColumn.Delete()
$WorkSheet.Range("AD1:AD1").EntireColumn.Delete()
$WorkSheet.Range("H1:Y1").EntireColumn.Delete()
$WorkSheet.Range("A1:C1").EntireColumn.Delete()


#Delete Blank Rows
[void]$Worksheet.Cells.Item(1).EntireRow.Delete()
[void]$Worksheet.Cells.Item(2).EntireRow.Delete()
[void]$Worksheet.Cells.Item(3).EntireRow.Delete()
[void]$Worksheet.Cells.Item(4).EntireRow.Delete()


$max = $WorkSheet.UsedRange.Rows.Count
for ($i = $max; $i -ge 0; $i--) {
    If ($WorkSheet.Cells.Item($i, 1).text -eq "") {
        $Range = $WorkSheet.Cells.Item($i, 1).EntireRow
        [void]$Range.Delete()
    } 
    Else {$i = 0}
} 

#Quit Excel
[void]$workbook.save() # Save file
[void]$workbook.close() # Close file
$Excel.Quit()