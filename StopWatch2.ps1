$StopWatch = [Diagnostics.Stopwatch]::StartNew()

$null = Read-Host -Prompt "Press ENTER as fast as you can!"

$StopWatch.Stop()
$StopWatch.ElapsedMilliseconds