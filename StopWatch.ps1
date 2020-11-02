

$Stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

$Timespan = [System.Timespan]$Stopwatch.ElapsedTicks
write-host $Timespan