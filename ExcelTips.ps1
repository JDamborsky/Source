#####################################################################
# Open New Excel Workbook / WorkSheet
$Excel = new-object -comobject excel.application
$ExcelWordBook = $Excel.Workbooks.Add()
$ExcelWorkSheet = $ExcelWordBook.Worksheets.Add()
$Excel.Visible = $true
 
#####################################################################
## Load Excel  file
$ExcelPath = 'C:\KM_Main.xlsx'
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$ExcelWordBook = $Excel.Workbooks.Open($ExcelPath)
$ExcelWorkSheet = $Excel.WorkSheets.item(&amp;quot;Sheet 1&amp;quot;)
 
#####################################################################
# Close connections to Excel
# set interactive to false so no save buttons are shown
$Excel.DisplayAlerts = $false
$Excel.ScreenUpdating = $false
$Excel.Visible = $false
$Excel.UserControl = $false
$Excel.Interactive = $false
## save the workbook
$Excel.Save()
## quit the workbook
$Excel.Quit()
## function to close all com objects
function Release-Ref ($ref) {
([System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) -gt 0)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
}
## close all object references
Release-Ref($ExcelWorkSheet)
Release-Ref($ExcelWordBook)
Release-Ref($Excel)
 
#####################################################################
# Change to a different Worksheet
$ExcelWorkSheet = $Excel.WorkSheets.item("Sheet 2")
 
#####################################################################
# Update / Insert / Delete Value in a Cell
$ExcelWorkSheet.Cells.Item(1,1).Value2 = "New Value"
 
#####################################################################
# Read Cell
$ExcelWorkSheet.Cells.Item(1,1).Text
 
#####################################################################
# Delete Row / Column
[void]$ExcelWorkSheet.Cells.Item(1,1).EntireColumn.Delete()
[void]$ExcelWorkSheet.Cells.Item(1,1).EntireRow.Delete()
 
#####################################################################
# Find Last Used Column or Row
$ExcelWorkSheet.UsedRange.columns.count
$ExcelWorkSheet.UsedRange.rows.count
 
#####################################################################
# Sorting
$table = $ExcelWorkSheet.ListObjects | where DisplayName -EQ "User_Table"
$table.Sort.SortFields.clear()
$table.Sort.SortFields.add( $table.Range.Columns.Item(1) )
$table.Sort.apply()
 
#####################################################################
# Clear all formatting on a sheet
$tableRange = $ExcelWorkSheet.UsedRange
$tableRange.ClearFormats()
 
#####################################################################
# Using Excel Table Styles&amp;lt;
## formatting from &amp;lt;a href="http://activelydirect.blogspot.co.uk/2011/03/write-excel-spreadsheets-fast-in.html"&amp;gt;http://activelydirect.blogspot.co.uk/2011/03/write-excel-spreadsheets-fast-in.html&amp;lt;/a&amp;gt;
$listObject = $ExcelWorkSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $ExcelWorkSheet.UsedRange, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null)
$listObject.Name = "User Table"
$listObject.TableStyle = "TableStyleLight10"
 
#####################################################################
# Auto-Sizing Columns / Rows
$ExcelWorkSheet.UsedRange.Columns.Autofit() | Out-Null
 
#####################################################################
# Formatting a Column
$ExcelWorkSheet.columns.item($formatcolNum).NumberFormat = "yyyy/mm/dd"
 
#####################################################################
# Formatting Text / Numbers Colours
# &lt;a href="http://theolddogscriptingblog.wordpress.com/2009/08/04/adding-color-to-excel-the-powershell-way/"&gt;http://theolddogscriptingblog.wordpress.com/2009/08/04/adding-color-to-excel-the-powershell-way/&lt;/a&gt;
 
#####################################################################
# Format Text / Numbers Bold
$ExcelWorkSheet.Cells.Item(1,1).Font.Bold=$True
 
#####################################################################
# Add Hyperlink to cell
$link = "http://www.microsoft.com/technet/scriptcenter"
$r = $ExcelWorkSheet.Range("A2")
[void]$ExcelWorkSheet.Hyperlinks.Add($r, $link)
 
#####################################################################
# Add Comment to Cell
$ExcelWorkSheet.Range("D2").AddComment("Autor Name: `rThis is my comment")
 
#####################################################################
# Add a Picture to a Comment
$image = "C:\test\Pictures\Kittys\gotyou.jpg"
$ExcelWorkSheet.Range("C1").AddComment()
$ExcelWorkSheet.Range("d3").Comment.Shape.Fill.UserPicture($image)
 
#####################################################################
# Fix Location and Size of comment
$ExcelWorkSheet.Range("D3").Comment.Shape.Left = 100
$ExcelWorkSheet.Range("D3").Comment.Shape.Top = 100
$ExcelWorkSheet.Range("D3").Comment.Shape.Width = 100
$ExcelWorkSheet.Range("D3").Comment.Shape.Height = 100
 
#####################################################################
# Making a Comment/s visible
$comments = $ExcelWorkSheet.comments
foreach ($c in $comments) {
$c.Visible = 1
}