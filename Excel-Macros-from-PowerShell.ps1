# file path to your XLA file with macros
$FilePath = "c:\test\file.xla"
# macro name to run
$Macro = "AddData"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$wb = $excel.Workbooks.Add($FilePath)
$excel.Run($Macro)