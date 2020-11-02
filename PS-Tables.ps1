PowerShell

$table = New-Object System.Data.DataTable

$c = $table.Columns.Add("Title","string") 
$c.ReadOnly = $True
$c.Unique = $True

$table.Columns.Add("ReleaseDate","datetime") | Out-Null
$table.Columns.Add("OpensIn","int32") | Out-Null
$table.Columns.Add("Comments","string") | Out-Null
$table.Columns.Add("Rating","string") | Out-Null
$table.Columns.add("Released","boolean") | Out-Null

foreach ($item in $data) {
    #define a new row object
    $r = $table.NewRow()
 
    #set the properties of each row from the data
    $r.Title = $item.Title
    $r.ReleaseDate = $item.ReleaseDate
    $r.OpensIn = ($r.ReleaseDate - (Get-Date)).TotalDays
    $r.Comments = $item.Comments
    $r.Rating = $item.Rating
    $r.Released = if ($r.OpensIn -lt 0) {
        $True
        }
        else {
        $False  
        }
    #add the row to the table
    $table.Rows.Add($r)
    } #foreach

$table.Rows.Add($r)


	
$m = $table.where({$_.title -eq 'Justice League'})