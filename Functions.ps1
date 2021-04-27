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

function ConvertCharsToAdsiSearcher
{
	param (
		[string]$AdsiSearchString,
		[Bool]$Reverse = $False)
	#   Adsisearcher does fail with some characters
	
	If ($Reverse -eq $False)
	{
		#   $AdsiSearchString = $AdsiSearchString -replace '\+', '+'
		$AdsiSearchString = $AdsiSearchString -replace '[#(]', '\28'
		$AdsiSearchString = $AdsiSearchString -replace '[#)]', '\29'
	}
	else
	{
		$AdsiSearchString = $AdsiSearchString.replace('\28', '(')
		$AdsiSearchString = $AdsiSearchString.replace('\29', ')')
	}
	
	$ReturnString = $AdsiSearchString
	Return $ReturnString
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

Function IsConnectetAdapterWired
{
    #    Check if connected networkadapter is Wired   (to use for prereq in WinPe)
    #$WirelessConnected = $null
    #$WiredConnected = $null
    #$VPNConnected = $null
    $ConnectionStatus = $null

        
    $WirelessAdapters =  Get-CimInstance -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter `
        'NdisPhysicalMediumType = 9'
    $WiredAdapters = Get-CimInstance -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter `
        "NdisPhysicalMediumType = 0 and `
        (NOT InstanceName like '%pangp%') and `
        (NOT InstanceName like '%cisco%') and `
        (NOT InstanceName like '%juniper%') and `
        (NOT InstanceName like '%vpn%') and `
        (NOT InstanceName like 'Hyper-V%') and `
        (NOT InstanceName like 'VMware%') and `
        (NOT InstanceName like 'VirtualBox Host-Only%')"
    $ConnectedAdapters =  Get-CimInstance -Class win32_NetworkAdapter -Filter `
        'NetConnectionStatus = 2'
    $VPNAdapters =  Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter `
        "Description like '%pangp%' `
        or Description like '%cisco%'  `
        or Description like '%juniper%' `
        or Description like '%vpn%'"

    Foreach($Adapter in $ConnectedAdapters) {
        If($WiredAdapters.InstanceName -contains $Adapter.Name)
        {
     #       $WiredConnected = $true
            $ConnectionStatus = "Wired"
        }
    }

    Foreach($Adapter in $ConnectedAdapters) {
        If($WirelessAdapters.InstanceName -contains $Adapter.Name)
        {
      #      $WirelessConnected = $true
            $ConnectionStatus = "Wireless"
        }
    }

    Foreach($Adapter in $ConnectedAdapters) {
        If($VPNAdapters.Index -contains $Adapter.DeviceID)
        {
       #     $VPNConnected = $true
            if ($ConnectionStatus -eq "Wired")
            {
                $ConnectionStatus = "Wired-VPN"
            }
            else 
            {
                $ConnectionStatus = "VPN"    
            }
            
        }
    }
    Return $ConnectionStatus
}

function ConvertTo-DataTable
	{
		<#
			.SYNOPSIS
				Converts objects into a DataTable.
		
			.DESCRIPTION
				Converts objects into a DataTable, which are used for DataBinding.
		
			.PARAMETER  InputObject
				The input to convert into a DataTable.
		
			.PARAMETER  Table
				The DataTable you wish to load the input into.
		
			.PARAMETER RetainColumns
				This switch tells the function to keep the DataTable's existing columns.
			
			.PARAMETER FilterCIMProperties
	            This switch removes CIM properties that start with an underline.
	        
	        .PARAMETER MatchColumns
				This switch force only listed properties to be included ( Use |  as delimiter)
		
			.EXAMPLE
				$DataTable = ConvertTo-DataTable -InputObject (Get-Process)
		#>
		[OutputType([System.Data.DataTable])]
		param (
			$InputObject,
			[ValidateNotNull()]
			[System.Data.DataTable]$Table,
			[switch]$RetainColumns,
			[switch]$FilterCIMProperties,
			[String]$MatchColumns)
		
		if ($null -eq $Table)
		{
			$Table = New-Object System.Data.DataTable
		}
		
		if ($null -eq $InputObject)
		{
			$Table.Clear()
			return @( ,$Table)
		}
		
		if ($InputObject -is [System.Data.DataTable])
		{
			$Table = $InputObject
		}
		elseif ($InputObject -is [System.Data.DataSet] -and $InputObject.Tables.Count -gt 0)
		{
			$Table = $InputObject.Tables[0]
		}
		else
		{
			if (-not $RetainColumns -or $Table.Columns.Count -eq 0)
			{
				#Clear out the Table Contents
				$Table.Clear()
				
				if ($null -eq $InputObject) { return } #Empty Data
				
				$object = $null
				#find the first non null value
				foreach ($item in $InputObject)
				{
					if ($null -ne $item)
					{
						$object = $item
						break
					}
				}
				
				if ($null -eq $object) { return } #All null then empty
				
				#Get all the properties in order to create the columns
				foreach ($prop in $object.PSObject.Get_Properties())
				{
					if ('RowError', 'RowState', 'Table', 'ItemArray', 'HasErrors' -contains $prop.Name)
					{
						continue
					}
					If ($PSBoundParameters.ContainsKey('MatchColumns'))
					{
						$MatchColumnsArray = $MatchColumns.split("|")
						if (!$MatchColumnsArray.contains($prop.Name))
						{
							continue
						}
					}
					if (-not $FilterCIMProperties -or -not $prop.Name.StartsWith('__')) #filter out CIM properties
					{
						
						#Get the type from the Definition string
						$type = $null
						
						if ($null -ne $prop.Value)
						{
							try { $type = $prop.Value.GetType() }
							catch { Out-Null }
						}
						
						if ($null -ne $type) # -and [System.Type]::GetTypeCode($type) -ne 'Object')
						{
							[void]$table.Columns.Add($prop.Name, $type)
						}
						else #Type info not found
						{
							[void]$table.Columns.Add($prop.Name)
						}
					}
				}
				
				if ($object -is [System.Data.DataRow])
				{
					foreach ($item in $InputObject)
					{
						#$Table.Rows.Add($item)
						$Table.ImportRow($item)
					}
					return @( ,$Table)
				}
			}
			else
			{
				$Table.Rows.Clear()
			}
			
			foreach ($item in $InputObject)
			{
				$row = $table.NewRow()
				
				if ($item)
				{
					foreach ($prop in $item.PSObject.Get_Properties())
					{
						if ($table.Columns.Contains($prop.Name))
						{
							$row.Item($prop.Name) = $prop.Value
						}
					}
				}
				[void]$table.Rows.Add($row)
			}
		}
		
		return @( ,$Table)
	}
	
	


function Get-AdsiPathForCurrentDomain
{
    $Root = [ADSI]"LDAP://RootDSE"
    $GetAdsiPathStr = 'LDAP://' + $Root.rootDomainNamingContext
    return $GetAdsiPathStr
}

function Get-MachineInfoFromAD
{
       param (
              [string]$ClientNameToFind)
       
       $IsObjectFound           = $false
       $operatingsystemversion  = ''
       $operatingsystem         = ''
       $lastLogon               = ''
       $LLDate                  = ''
       $MemberOf                = ''
       $MachineLdap             = ''
       
     
       # Build for ADSI Query
       $AdsiPathForDomain 		= Get-AdsiPathForCurrentDomain
       $searcher 				= [adsisearcher]"(&(objectCategory=computer)(cn=$ClientNameToFind))"
       $searcher.searchRoot 	= [ADSI]$AdsiPathForDomain
       $searcher.PropertiesToLoad.Add('operatingsystemversion')
       $searcher.PropertiesToLoad.Add('operatingsystem')
       $searcher.PropertiesToLoad.Add('lastlogon')
       $searcher.PropertiesToLoad.Add('memberof')
       $searcher.PropertiesToLoad.Add('distinguishedname')
       $searcher.PropertiesToLoad.Add('useraccountcontrol')
       $searcher.PropertiesToLoad.Add('pwdLastSet')
       try
       {
              $MachineItemList = $searcher.FindAll()
       }
       catch
       {
              $IsObjectFound = $false
              Write-Host "Not Found in AD"
       }
       
       foreach ($MachineItem in $MachineItemList)
       {
            $IsObjectFound              = $true
            $IsObjectEnabled            = $true
            $MachineLdap                = $MachineItem.Path
            $operatingsystemversion     = $MachineItem.Properties.Item('operatingsystemversion')
            $operatingsystem            = $MachineItem.Properties.Item('operatingsystem')
            
            $LL = $MachineItem.Properties.Item("lastLogon")[0]            
            If (-Not $LL) { $LL = 0 }
            $LLDate         = [DateTime]$LL
            $lastLogon      = $LLDate.AddYears(1600).ToLocalTime()
            $LL             = $MachineItem.Properties.Item("pwdLastSet")[0]            
            If (-Not $LL) { $LL = 0 }
            $LLDate         = [DateTime]$LL
            $pwdLastSet     = $LLDate.AddYears(1600).ToLocalTime()
            
            if ([string]$MachineItem.properties.useraccountcontrol -band 2)
            {
                    $IsObjectEnabled    = $False
                    $IsObjectEnabledbit = 0
            }
            else
            {
                    $IsObjectEnabled    = $true
                    $IsObjectEnabledbit = 1
            }
            
            $MemberOfArray = $MachineItem.Properties.Item('memberof')
            
       }
      
       # Build a string of all Group-Names
       $memberof = ''
       foreach ($MemberOfItem in $MemberOfArray)
       {
              $MemberOfItemArray    = $MemberOfItem -split ','
              $MemberOfGroup        = $MemberOfItemArray[0] -replace 'CN='
              $memberof             += "$MemberOfGroup;"
       }
       
       if ($memberof.Length -gt 0)
       {
              $memberof = $memberof.Substring(0, $memberof.Length - 1)
       }
    
       # Build a PsCustomObject to return
	   
       $tmpHashtable = @{
              MachineName 				= "$ClientNameToFind"
			  IsObjectFound             = $IsObjectFound
              MachineLdap               = "$MachineLdap"
              operatingsystem           = "$operatingsystem"
              operatingsystemversion    = "$operatingsystemversion"
              lastLogon                 = "$lastLogon"
              #memberofList              = "$memberof"
              #memberofArray             = "$MemberOfArray" 
              IsObjectEnabled           = $IsObjectEnabled
              pwdLastSet                = "$pwdLastSet"
       }


	   #  Convert PsCustomObject to DataTable
       $ResultTable = New-Object System.Data.DataTable
       #Headdings
       foreach ($Item in $tmpHashtable.keys)
       {
        $col = New-Object System.Data.DataColumn("$Item")
        $ResultTable.columns.Add($col)        
       }
	   #Values
       $row = $ResultTable.NewRow()
       $ColCounter = 0
       foreach ($ValItem in $tmpHashtable.Values)
       { 
          $RowName = $ResultTable.columns[$ColCounter].ColumnName
          $row["$RowName"] = $ValItem
          $ColCounter++
       } 
       $ResultTable.rows.Add($row)


       # Choose output ......
	   #Return [pscustomobject]$tmpHashtable
    	Return [System.Data.DataTable]$ResultTable

       
}





Function Get-AdNetInfoForIP
	{
		[CmdletBinding()]
		param (
			[Parameter()]
			[IPAddress]$Ip4Adress
		)

		$FoundAdNetId   	= ""
        $DomainName      	= $([adsi] "LDAP://RootDSE").Get("rootDomainNamingContext")
        $DomainShortName 	= $DomainName.Split(",")[0].Replace('DC=','')
		#$sitesDN        	= "LDAP://CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
		$subnetsDN      	= "LDAP://CN=Subnets,CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
		
		foreach ($subnet in $([adsi]$subnetsDN).psbase.children)
		{
           	$CurrNetAdr     = ([IPAddress](($subnet.cn -split "/")[0]))
			$CurrAdSn       = ([IPAddress]"$([system.convert]::ToInt64(("1" * [int](($subnet.cn -split "/")[1])).PadRight(32, "0"), 2))")
			if ((([IPAddress]$Ip4Adress).Address -band ([IPAddress]$CurrAdSn).Address) -eq ([IPAddress]$CurrNetAdr).Address)
			{
				$FoundAdNetId   = $($subnet.cn)
				
				$site = [adsi] "LDAP://$($subnet.siteObject)"
				if ($site.cn -ne $null)
				{
					$siteName   		= ([string]$site.cn).toUpper()
					$siteDescription	= ([string]$site.description)					
  				}
				
				$SubNetDescription  = $subnet.description[0]
				$SubNetLocation     = $subnet.Location[0]
				$AdSiteForAdress    = @{
					ip                  = "$Ip4Adress"
					sn                  = "$CurrAdSn"
					AdCidr              = "$FoundAdNetId"
					AdSiteName          = "$siteName"
					AdSiteDesciption    = "$siteDescription"
					SubNetDescription   = "$SubNetDescription"
					SubNetLocation      = "$SubNetLocation"
                    DomainName          = "$DomainShortName"
					Isfound             = $True
				}
			#	Break				
			}
		}
		if ($FoundAdNetId -eq "")
		{
			$AdSiteForAdress = @{
				ip                  = "$Ip4Adress"
				sn                  = ""
				AdCidr              = ""
				AdSiteName          = ""
				AdSiteDesciption    = ""
				SubNetDescription   = ""
				SubNetLocation      = ""
                DomainName          = ""
				Isfound             = $False
			}
		}
		$FoundAdNetIdObject = [pscustomobject]$AdSiteForAdress
		Return $FoundAdNetIdObject		
	}

