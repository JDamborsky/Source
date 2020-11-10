#  Script to test some consepts around jobs running in parallel
#  Jarle@Damborsky.com      10 Nov. -20


# Clean up variables in memory
Remove-Variable * -ErrorAction SilentlyContinue
Remove-Variable * -scope 'Global' -ErrorAction SilentlyContinue
Get-Job | Remove-Job

# Function to return LDAP Searchroot from domain, also trusted domains
function GetLdapFromDomainShortName
{
	param (
		[string]$DomainShortnameToFind
	)
	$DomainArray = @()
	
	$CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
	$DomainArray += $CurrentDomain
	
	$TrustedDomainsList = $CurrentDomain.GetAllTrustRelationships()
	foreach ($TrustedDomain in $TrustedDomainsList)
	{
		$DomainArray += $TrustedDomain.Targetname
	}
	
	$Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
	$TrustedForestList = $Forest.GetAllTrustRelationships()
	foreach ($TrustedForest in $TrustedForestList)
	{
		$DomainArray += $TrustedForest.Targetname
	}
	
	foreach ($DomainItems in $DomainArray)
	{
		if ($DomainItems -like "*$DomainShortnameToFind*")
		{
			[String]$DomainName = $DomainItems.tostring()
			$LdapStr = "LDAP://"
			$DomainNameParts = $DomainName -split '\.'
			foreach ($DomainNamePart in $DomainNameParts)
			{
				$LdapStr = $LdapStr + "DC=$DomainNamePart,"
			}
			$LdapStr = $LdapStr.Substring(0, $LdapStr.length - 1)
		}
	}
	
	Return $LdapStr
}



$UserinfoFromDomainBlock = {
	param
	(
		[parameter(Mandatory = $true)]
		[String]$UsernameToFind,
        [parameter(Mandatory = $true)]
		[String]$LdapSearchRoot
	)
	
	$IsObjectFound	= $false
	
	$AdsiSearcher = [adsisearcher]"(&(objectCategory=Person)(objectClass=user)(sAMAccountName=$UsernameToFind))"
	$AdsiSearcher.PropertiesToLoad.AddRange(('name'))
	$AdsiSearcher.searchRoot = [ADSI]$LdapSearchRoot
	$AdsiSearchList = $AdsiSearcher.FindAll()
	
	foreach ($UserItem in $AdsiSearchList)
	{
		$IsObjectFound		=	$true
		$User_Name = $UserItem.Properties.Item('name')[0]
	}
	
	$tmpHashtable = @{
		ObjectFound = $IsObjectFound
		name 		= "$User_Name"
	}
	$ReturnObject = [pscustomobject]$tmpHashtable
	#Start-Sleep 10
	Return $ReturnObject
	
}

$MachinInfoFromDomainBlock = {
	param
	(
		[parameter(Mandatory = $true)]
		[String]$MachineNameToFind,		
        [parameter(Mandatory = $true)]
		[String]$LdapSearchRoot
	)
	
	$IsObjectFound	= $false
	If ($MachineNameToFind -notlike "*$") {$MachineNameToFind =  $MachineNameToFind + "$"}
	$AdsiSearcher = [adsisearcher]"(&(objectCategory=computer)(objectClass=computer)(sAMAccountName=$MachineNameToFind))"	
	$AdsiSearcher.searchRoot = [ADSI]$LdapSearchRoot
	$AdsiSearchList = $AdsiSearcher.FindOne()
	
	foreach ($MachineItem in $AdsiSearchList.properties)
	{
		$IsObjectFound		=	$true
		$Machine_Name 						= $MachineItem.Item('name')
		$Machine_distinguishedName 			= $MachineItem.Item('distinguishedname')
		$Machine_dNSHostName 				= $MachineItem.Item('dnshostname')		
		$Machine_lastLogon 					= $MachineItem.Item('lastlogon')
		$Machine_operatingSystem 			= $MachineItem.Item('operatingsystem')		
		$Machine_operatingSystemVersion 	= $MachineItem.Item('operatingsystemversion')
		$Machine_pwdLastSet 				= $MachineItem.Item('pwdlastset')

	}
	
	$tmpHashtable = @{
		ObjectFound 				= $IsObjectFound
		Cli_Name 					= "$Machine_Name"
		Cli_distinguishedName 		= "$Machine_distinguishedName"
		Cli_dNSHostName 			= "$Machine_dNSHostName"
		Cli_lastLogon 				= "$Machine_lastLogon"
		Cli_operatingSystem 		= "$Machine_operatingSystem"
		Cli_operatingSystemVersion 	= "$Machine_operatingSystemVersion"
		Cli_pwdLastSet 				= "$Machine_pwdLastSet"
	}
	$ReturnObject = [pscustomobject]$tmpHashtable
	#Start-Sleep 10
	Return $ReturnObject
	
}


$ComputerInfoBlock = {
	$TmpComputerinfo = Get-ComputerInfo
	return $TmpComputerinfo
}


Function Init-PsJob
{
	[CmdletBinding()]
	param (		
		[Management.Automation.ScriptBlock]$BlockName,
        [String]$JobName,
        [String]$VariableName,
        [Collections.ArrayList]$Argumentlist=""
    )
    if ($Argumentlist -eq "")
    {
        Start-Job -Scriptblock $BlockName -Name $JobName | Out-Null
    }
    else {
        Start-Job -Scriptblock $BlockName -Name $JobName -ArgumentList $Argumentlist | Out-Null
    }
    $Global:JobArray += @([PsCustomObject]@{ BlockName = "$BlockName"; JobName = "$JobName"; VarName = $VariableName })
    New-Variable -Name "$VariableName" -Value "" -Scope "Global"
}

# First job
$JobParameterArray          = [Collections.ArrayList] @()
Init-PsJob -BlockName $ComputerInfoBlock -JobName 'GetComputerInfo' -VariableName "GetComputerInfoResult" -ArgumentList $JobParameterArray


# Second job, Userinfo from AD
$CurrentUser                = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$Global:UserName            = $CurrentUser.split("\")[1]
$Global:UserDomainShort     = $CurrentUser.split("\")[0]
$Global:UserDomainLdap      = GetLdapFromDomainShortName $Global:UserDomainShort

$JobParameterArray          = [Collections.ArrayList] @($Global:UserName, $Global:UserDomainLdap)
Init-PsJob -BlockName $UserinfoFromDomainBlock -JobName 'GetUserDomainInfo' -VariableName "GetUserDomainInfoResult" -ArgumentList $JobParameterArray | Out-Null

# Third Job, Machineinfo from AD
$Global:MachineName         = (Get-WmiObject Win32_ComputerSystem).Name
$Global:MachineDomainFQDN   = (Get-WmiObject Win32_ComputerSystem).Domain
$Global:MachineDomainLdap   = GetLdapFromDomainShortName $Global:MachineDomainShort

$JobParameterArray          = [Collections.ArrayList] @($Global:MachineName, $Global:MachineDomainLdap)
Init-PsJob -BlockName $MachinInfoFromDomainBlock -JobName 'GetMachineDomainInfo' -VariableName "GetMachineDomainInfoResult" -ArgumentList $JobParameterArray | Out-Null


# All jobs are started, check if they are finished
While (($Global:GetComputerInfoResult -eq "") -or ($GetUserDomainInfoResult -eq "") -or ($GetMachineDomainInfoResult -eq "")) 
{
    foreach ($JobItem in $Global:JobArray)
    {
        if (($(Get-Variable -Name "$($JobItem.VariableName)" -ValueOnly) -eq "") -and ((get-job -Name "$($JobItem.JobName)" -ChildJobState Running).State -eq 'Completed'))
        {
          
            switch ($JobItem.JobName) {
                "GetComputerInfo"       {   $Global:GetComputerInfoResult        = Receive-Job -Name "$($JobItem.JobName)"
                                            $Global:GetComputerInfoResult       | Format-List                                     }
                "GetUserDomainInfo"     {   $Global:GetUserDomainInfoResult      = Receive-Job -Name "$($JobItem.JobName)"
                                            $Global:GetUserDomainInfoResult     | Format-List                                     }
                "GetMachineDomainInfo"  {   $Global:GetMachineDomainInfoResult   = Receive-Job -Name "$($JobItem.JobName)"
                                            $Global:GetMachineDomainInfoResult  | Format-List                                     }
                Default {}
            }           
        }
    }

    Start-Sleep 1
}





