function GetAdsiPathForCurrentDomain
{
	$Root = [ADSI]"LDAP://RootDSE"
	$GetAdsiPathStr = 'LDAP://' + $($Root.rootDomainNamingContext)
	return $GetAdsiPathStr
}

Function GetGroupsUser
{
	param (
		[string]$UserNameStr
	)
	
	if ($UserNameStr.contains("\") -eq $True)
	{
		$UserNameArray = $UserNameStr.split("\")
		$UserName = $UserNameArray[1]
		$Domain = $UserNameArray[0]
	}
	else
	{
		$Domain = $env:USERDOMAIN 
		$UserName = $UserNameStr
	}
	
	$GroupArray = @()
	$Searcher = [adsisearcher]"(&(objectCategory=person)(objectClass=User)(sAMAccountName=$UserName))"
	$searcher.searchRoot = [ADSI]$Global:AdsiPathString
	$searcher.PropertiesToLoad.Add('memberof')
	$AdSearchList = $searcher.FindAll()
	foreach ($AdSearchItem in $AdSearchList.properties.memberof)
	{
	
		$GroupName = GetGroupnameFromDn $AdSearchItem $Global:AdsiPathString
	
		$GroupArray += @([PsCustomObject]@{ Name = $GroupName; DN = $AdSearchItem; MemberType = 'Direct' })
	}
	
	$GroupArray = $GroupArray | Sort-Object DN -Unique
	Return $GroupArray
}


Function GetGroupnameFromDn
{
	param (
		$GroupDN,
		$LdapPathStr)
	
	$GroupnameFound = $false
	$RetryCount = 0
	$DumpToTable = $null
	$Groupname = $null
	
	while ((!$GroupnameFound) -and ($RetryCount -lt 2))
	{
		$RetryCount++
		if ($RetryCount -eq 1)
		{
			$GroupDN = ConvertCharsToAdsiSearcher $GroupDN
		}
		
		$searcher = [adsisearcher]"(&(objectClass=Group)(distinguishedName=$GroupDN))"
		$searcher.searchRoot = [adsi]$LdapPathStr
		$Searcher.SizeLimit = 12
		$Searcher.PageSize = 11
		$searcher.PropertiesToLoad.AddRange(('name', 'samaccountname'))
		try
		{
			$DumpToTable = $Searcher.FindAll()
		}
		catch
		{
			$DumpToTable = 'fail'
		}
		
		if ($DumpToTable -ne 'fail')
		{
			foreach ($groupItem in $DumpToTable)
			{
				$Groupname = $groupItem.Properties.name
				$GroupnameFound = $true
			}
		}
	}
	
	if ($RetryCount -gt 1)
	{
		write-host " Group not found ->$GroupDN        "
	}
	
	return $Groupname
}


function ListMemberofForUser
{
	$UserDirectMemberOf = GetGroupsUser 'jarle'
	foreach ($GroupItem in $UserDirectMemberOf)
	{
		Write-Host $GroupItem.name
		
    }

    $UserDirectMemberOf | Export-Clixml  -Path "C:\data\gruppeliste.xml"
    
}
