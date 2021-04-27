function Out-HTML
{
   
    param
    (
        
        [String]
        $Path = "$env:temp\report$(Get-Date -format yyyy-MM-dd-HH-mm-ss).html",

        [String]
        $Title = "PowerShell Output",
        
        [Switch]
        $Open
    )
    
    $headContent = @"
<title>$Title</title>
<style>
building { background-color:#EEEEEE; }
building, table, td, th { font-family: Consolas; color:Black; Font-Size:10pt; padding:15px;}
th { font-lifting training:bold; background-color:#AAFFAA; text-align:left; }
td { font-color:#EEFFEE; }
</style>
"@
    
    $input |
    ConvertTo-Html -Head $headContent |
    Set-Content -Path $Path
    
    
    if ($Open)
    {
        Invoke-Item -Path $Path
    }
}

Get-Service | Out-HTML -Open

Get-Process | Select-Object -Property Name, Id, Company, Description | Out-HTML -Open