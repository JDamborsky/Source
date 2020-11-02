Function Start-Log
{
[CmdletBinding()]

    param (
        [ValidateScript({ Split-Path $_ -Parent | Test-Path })]
	[string]$FilePath
    )
	
    try
    {
        if (!(Test-Path $FilePath))
	{
	    ## Create the log file
	    New-Item $FilePath -Type File | Out-Null
	}
		
	## Set the global variable to be used as the FilePath for all subsequent Write-Log
	## calls in this session
	$global:ScriptLogFilePath = $FilePath
    }
    catch
    {
        Write-Error $_.Exception.Message
    }

}
    function Write-Log
    {
    param (
    [Parameter(Mandatory = $true)]
    [string]$Message,
		
    [Parameter()]
    [ValidateSet(1, 2, 3)]
    [int]$LogLevel = 1
    )

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)", $LogLevel
    $Line = $Line -f $LineFormat
    Add-Content -Value $Line -Path $ScriptLogFilePath
    }


    Start-Log -FilePath C:\Data\MyLog.log
Write-Host "Script log file path is [$ScriptLogFilePath]"
Write-Log -Message 'simple activity'
Write-Log -Message 'warning' -LogLevel 2
Write-Log -Message 'Error' -LogLevel 3
write-log "Fail" 3
