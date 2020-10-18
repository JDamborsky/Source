# For use in Task Sequence to move current machine to new OU

param (
        [String]$NewOuName
    )
Function Logg
{
    param (
        [String]$Loggstring )
        $Loggstring = (get-date).ToString('G') +": " + $Loggstring
        write-host $Loggstring
}

function GetCurrentOuLdapForMachine
{
    $CurrentComputername    = $(Get-WmiObject Win32_Computersystem).name
    $CurrentComputerLdap    = ([adsisearcher]"samaccountname=$CurrentComputername$").findone()
    $CurrentComputerLdapStr = $( $CurrentComputerLdap.path) -replace 'LDAP://'
    $TmpArray  = $CurrentComputerLdapStr -split ","
    $CurrentComputerLdapStr = $CurrentComputerLdapStr -replace "$($TmpArray[0]),"
    return $CurrentComputerLdapStr
}

Function MoveTo
{
    param (
        [String]$NewOuName
    )

    if ($NewOuName.contains("LDAP://")) { $NewOuName = $NewOuName -replace 'LDAP://' }
    $SysInfo = New-Object -ComObject "ADSystemInfo"
    $ComputerDN = $SysInfo.GetType().InvokeMember("ComputerName", "GetProperty", $Null, $SysInfo, $Null)
    logg "Current location of this PC : $ComputerDN"
    $Computer = [ADSI]"LDAP://$ComputerDN"
    logg "Moving PC to : $NewOuName"
    try
    { 
        $Computer.psbase.MoveTo([ADSI]"LDAP://$($NewOuName)") 
    }
    catch {
        $ResultStr =  $_
        logg "ResultStr = $ResultStr"
    }

    start-sleep -Seconds 10

    if (GetCurrentOuLdapForMachine -eq $NewOuName)
    {    $returnValue = "MoveToOU = Success"   }
    else 
    {    $returnValue = "MoveToOU = Failed"   }
    Return $returnValue
}

$Result = MoveTo $NewOuName
Logg "Result = $Result"

