$Path = “C:\Users\FLast\Desktop”

$FixedFormat = “Microsoft.Office.Interop.Excel.xlFixedFormatType” -as [type]

$ExcelFiles = Get-ChildItem -Path $Path -include *.xls, *.xlsx -recurse

$Excel = New-Object -ComObject excel.application
$Excel.visible = $false

foreach($ExcelFile in $ExcelFiles)
{
 $PDFFile = Join-Path -Path $Path -ChildPath ($ExcelFile.BaseName + “.pdf”)

 $WorkBook = $Excel.WorkBooks.Open($ExcelFile.FullName, 3)
 
 $WorkSheet = $WorkBook.WorkSheets.Item(1)
  
 $WorkSheet.PageSetup.CenterHeader = ""
 $WorkSheet.PageSetup.CenterFooter = ""

 #$WorkSheet.Range("1:1").EntireRow.Delete()
 #$WorkSheet.Range("Q:CA").EntireColumn.Delete()
 #$WorkSheet.Range("A:O").EntireColumn.Delete()

 $WorkBook.Saved = $True
 $WorkBook.ExportAsFixedFormat($FixedFormat::xlTypePDF, $PDFFile)
 
 $Excel.WorkBooks.close()
}

$Excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
SPPS -n Excel