#  Script to test some consepts around jobs running in parallel
#  Jarle@Damborsky.com      10 Nov. -20

# Clean up variables in memory
Remove-Variable * -ErrorAction SilentlyContinue
Remove-Variable * -scope 'Global' -ErrorAction SilentlyContinue

$Global:Logging = $true
Start-Log
write-log "Start"  1

Get-Job | Remove-Job

Function Start-Log
{
	if ($Global:Logging -eq $True)
	{
		$ScriptName         = Split-Path @(Get-PSCallStack)[0].invocationinfo.PSCommandpath -Leaf
		$ScriptName         = $ScriptName.Replace('.ps1','')
		$LoggFileFolder     = "$([Environment]::GetEnvironmentVariable("Temp", [EnvironmentVariableTarget]::User))\logs"    
		if (!(Test-Path $LoggFileFolder))
		{
			New-Item -ItemType directory -Path $LoggFileFolder
		}
		$FilePath = "$LoggFileFolder\$ScriptName-$DateStr.log"

		try
		{
			if (!(Test-Path $FilePath))	{	    New-Item $FilePath -Type File | Out-Null	}
			$global:ScriptLogFilePath = $FilePath
		}
		catch
		{	Write-Error $_.Exception.Message    }
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
	if ($Global:Logging -eq $True)
	{
		$TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
		$Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
		$LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($MyInvocation.ScriptName | Split-Path -Leaf):$($MyInvocation.ScriptLineNumber)", $LogLevel
		$Line = $Line -f $LineFormat
		Add-Content -Value $Line -Path $global:ScriptLogFilePath 
		#$Line | Out-File $($global:ScriptLogFilePath) -Append 
	}
}

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

$MachinInfoFromDomainBlock = {
    param
    (
        [parameter(Mandatory = $true)]
        [String]$MachineNameToFind,     
        [parameter(Mandatory = $true)]
        [String]$LdapSearchRoot
    )
    
    $IsObjectFound  = $false
    If ($MachineNameToFind -notlike "*$") {$MachineNameToFind =  $MachineNameToFind + "$"}
    $AdsiSearcher = [adsisearcher]"(&(objectCategory=computer)(objectClass=computer)(sAMAccountName=$MachineNameToFind))"   
    $AdsiSearcher.searchRoot = [ADSI]$LdapSearchRoot
    $AdsiSearchList = $AdsiSearcher.FindOne()
    
    foreach ($MachineItem in $AdsiSearchList.properties)
    {
        $IsObjectFound      =   $true
        $Machine_Name                       = $MachineItem.Item('name')
        $Machine_distinguishedName          = $MachineItem.Item('distinguishedname')
        $Machine_dNSHostName                = $MachineItem.Item('dnshostname')      
        $Machine_lastLogon                  = $MachineItem.Item('lastlogon')
        $Machine_operatingSystem            = $MachineItem.Item('operatingsystem')      
        $Machine_operatingSystemVersion     = $MachineItem.Item('operatingsystemversion')
        $Machine_pwdLastSet                 = $MachineItem.Item('pwdlastset')

    }
    
    $tmpHashtable = @{
        ObjectFound                 = $IsObjectFound
        Cli_Name                    = "$Machine_Name"
        Cli_distinguishedName       = "$Machine_distinguishedName"
        Cli_dNSHostName             = "$Machine_dNSHostName"
        Cli_lastLogon               = "$Machine_lastLogon"
        Cli_operatingSystem         = "$Machine_operatingSystem"
        Cli_operatingSystemVersion  = "$Machine_operatingSystemVersion"
        Cli_pwdLastSet              = "$Machine_pwdLastSet"
    }
    $ReturnObject = [pscustomobject]$tmpHashtable
    #Start-Sleep 10
    Return $ReturnObject
    
}

$MachinInfoFromWmiBlock = {
    $MachineHwArray += @([PsCustomObject]@{ Description = "Makin-navn:"; Value = $((Get-WmiObject Win32_ComputerSystem).Name) })
    $MachineHwArray += @([PsCustomObject]@{ Description = "Produsent:"; Value = $((Get-WmiObject Win32_ComputerSystem).Manufacturer) })
    $MachineHwArray += @([PsCustomObject]@{ Description = "Modell:"; Value = $((Get-WmiObject Win32_ComputerSystem).Model) })
    
    $MemInGb = Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | Foreach { "{0:N2}" -f ([math]::round(($_.Sum / 1GB), 2)) }
    $MachineHwArray += @([PsCustomObject]@{ Description = "Minne (GigaByte):"; Value = $MemInGb })
    
    $MachineHwArray += @([PsCustomObject]@{ Description = "Serienr:"; Value = $((Get-WmiObject Win32_Bios).SerialNumber) })
    $MachineHwArray += @([PsCustomObject]@{ Description = "Bios Versjon:"; Value = $((Get-WmiObject Win32_Bios).SMBIOSBIOSVersion) })
    
    $HddList = Get-WmiObject Win32_DiskDrive
    foreach ($HddItem in $HddList)
    {
        $MachineHwArray += @([PsCustomObject]@{ Description = "Disk:"; Value = $($HddItem.model) })
    }
    
    $VideoControllers = Get-WmiObject Win32_VideoController
    foreach ($Controller in $VideoControllers)
    {
        $MachineHwArray += @([PsCustomObject]@{ Description = "Skjermkort Modell:"; Value = $($Controller.Caption) })
        #write-host $Controller.Caption
    }
    
    
    #$MachineHwArray += @([PsCustomObject]@{ Description = "Skjermkort Minne (Gb) :"; Value = [math]::round($((Get-WmiObject Win32_VideoController).AdapterRAM) /1Gb, 0) })
    $MonitorList = Get-WmiObject Win32_DesktopMonitor
    foreach ($MonitorItem in $MonitorList)
    {
        $MonitorDescription = "$($MonitorItem.MonitorManufacturer) -> $($MonitorItem.Description)"
        $MachineHwArray += @([PsCustomObject]@{ Description = "Monitor:"; Value = $MonitorDescription })
    }
    
    #[math]::round($((Get-WmiObject Win32_VideoController).CurrentNumberOfColors) /1Mb, 0)
    [math]::round($($LogicalDiskItem.FreeSpace) /1Gb, 0)
    
    
    
    $OkFan = 0
    $FailFan = 0
    $FanList = Get-WmiObject Win32_Fan
    foreach ($FanItem in $FanList)
    {
        if ($($FanItem.Status) -eq "OK")
        { $OkFan++ }
        else
        { $FailFan++ }
    }
    $MachineHwArray += @([PsCustomObject]@{ Description = "Vifter:"; Value = "$OkFan er OK -- $FailFan har feilet" })
    
    $LogicalDiskList = Get-WmiObject Win32_LogicalDisk
    foreach ($LogicalDiskItem in $LogicalDiskList)
    {
        if ($($LogicalDiskItem.DeviceID) -eq "C:")
        { $MachineHwArray += @([PsCustomObject]@{ Description = "Ledig plass på C: (Gb):"; Value = [math]::round($($LogicalDiskItem.FreeSpace) /1Gb, 1) }) }
    }
    
    
    $UsbDeviceList = Get-WmiObject Win32_USBDevice
    foreach ($UsbDeviceItem in $UsbDeviceList)
    {
        if ($UsbDeviceItem.Description -eq "Disk drive")
        {
            $MachineHwArray += @([PsCustomObject]@{ Description = "USB Disk:"; Value = $UsbDeviceItem.Caption })
        }
        
    }
    
    return $MachineHwArray
}


$ComputerInfoBlock = {
    $TmpComputerinfo = Get-ComputerInfo
    return $TmpComputerinfo
}


$AdSubnetInfoBlock = {
	[CmdletBinding()]
	param (
		[Parameter()]
		[IPAddress]$Ip4Adress
	)
	
	$FoundAdNetId = ""
	$sitesDN = "LDAP://CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
	$subnetsDN = "LDAP://CN=Subnets,CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
	
	foreach ($subnet in $([adsi]$subnetsDN).psbase.children)
	{
		$CurrNetAdr = ([IPAddress](($subnet.cn -split "/")[0]))
		$CurrAdSn = ([IPAddress]"$([system.convert]::ToInt64(("1" * [int](($subnet.cn -split "/")[1])).PadRight(32, "0"), 2))")
		if ((([IPAddress]$Ip4Adress).Address -band ([IPAddress]$CurrAdSn).Address) -eq ([IPAddress]$CurrNetAdr).Address)
		{
			$FoundAdNetId = $($subnet.cn)
			
			$site = [adsi] "LDAP://$($subnet.siteObject)"
			if ($site.cn -ne $null)
			{
				$siteName = ([string]$site.cn).toUpper()
			}
			
			$SubNetDescription = $subnet.description[0]
			$SubNetLocation = $subnet.Location[0]
			$AdSiteForAdress = @{
				ip = "$Ip4Adress"
				sn = "$CurrAdSn"
				AdCidr = "$FoundAdNetId"
				AdSiteName = "$siteName"
				SubNetDescription = "$SubNetDescription"
				SubNetLocation = "$SubNetLocation"
				Isfound = $True
			}
			Break
			
		}
	}
	if ($FoundAdNetId -eq "")
	{
		$AdSiteForAdress = @{
			ip = "$Ip4Adress"
			sn = ""
			AdCidr = ""
			AdSiteName = ""
			SubNetDescription = ""
			SubNetLocation = ""
			Isfound = $False
		}
	}
	$FoundAdNetIdObject = [pscustomobject]$AdSiteForAdress
	Return $FoundAdNetIdObject
	
}

$CurrentNetworkBlock = {
	$NetworkCategory_Private_Found 	= $False
	$NetworkCategory_Public_Found 	= $False
	$NetworkCategory_Domain_Found 	= $False
	$IsOnVPN 						= $False
	$VpnIpAdress					= ""
	$VpnNetProfileName				= ""

	$NetProfilesList = Get-NetConnectionProfile
	foreach ($NetProfileItem in $NetProfilesList)
	{
		switch ($($NetProfileItem.NetworkCategory))
		{
			"Private"                { $NetworkCategory_Private_Found = $true }
			"Public"                 { $NetworkCategory_Public_Found = $true }
			"DomainAuthenticated"    { $NetworkCategory_Domain_Found = $true }
		}
		
		if ((($NetProfileItem.NetworkCategory -eq 'Private') -or ($NetProfileItem.NetworkCategory -eq 'Public')) -and (($NetProfileItem.IPv4Connectivity -eq 'LocalNetwork') -or ($NetProfileItem.IPv4Connectivity -eq 'Internet')))
		{
			
			#write-host  $(Get-NetIPAddress -InterfaceAlias $($NetProfileItem.InterfaceAlias) -AddressFamily IPv4).IPv4Address
			#Get-NetIPConfiguration -InterfaceAlias $($NetProfileItem.InterfaceAlias)  -AllCompartments       
		}
		
		if ($NetProfileItem.NetworkCategory -eq 'DomainAuthenticated')
		{
			$VpnIpAdress = $(Get-NetIPAddress -InterfaceAlias $($NetProfileItem.InterfaceAlias) -AddressFamily IPv4).IPv4Address
			
			$VpnNetProfileName = $(Get-NetIPConfiguration -AllCompartments -InterfaceAlias $($NetProfileItem.InterfaceAlias)).NetProfile.Name
			#Get-NetIPConfiguration -AllCompartments -InterfaceAlias $($NetProfileItem.InterfaceAlias)
		}
	}
	
	if (($NetworkCategory_Private_Found -or $NetworkCategory_Public_Found) -and $NetworkCategory_Domain_Found)
	{
		$IsOnVPN = $true
	}

	$CurrNetworkProfile = @{
		IsOnVPN 				= $IsOnVPN
		Private 				= $NetworkCategory_Private_Found
		Public 					= $NetworkCategory_Public_Found
		DomainAuthenticated 	= $NetworkCategory_Domain_Found
		VpnIpAdress 			= $VpnIpAdress
		VpnNetProfileName 		= $VpnNetProfileName
	}
	
	$DetectedNetworkProfile = [pscustomobject]$CurrNetworkProfile
	Return $DetectedNetworkProfile
	
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

#$Global:MachineDomainLdap   = GetLdapFromDomainShortName $Global:MachineDomainShort
$Global:MachineDomainLdap   = GetLdapFromDomainShortName $Global:MachineDomainFQDN

$JobParameterArray          = [Collections.ArrayList] @($Global:MachineName, $Global:MachineDomainLdap)
Init-PsJob -BlockName $MachinInfoFromDomainBlock -JobName 'GetMachineDomainInfo' -VariableName "GetMachineDomainInfoResult" -ArgumentList $JobParameterArray | Out-Null

# Job, Machineinfo from WMI etc..
$JobParameterArray          = [Collections.ArrayList] @()
Init-PsJob -BlockName $MachinInfoFromWmiBlock -JobName 'GetMachineWmiInfo' -VariableName "GetMachineWmiInfoResult" -ArgumentList $JobParameterArray | Out-Null

$Global:ipaddress = ([System.Net.DNS]::GetHostAddresses($labelMachineName.Text) | Where-Object { $_.AddressFamily -eq "InterNetwork" } | select-object IPAddressToString)[0].IPAddressToString
$JobParameterArray          = [Collections.ArrayList] @($Global:ipaddress)
Init-PsJob -BlockName $AdSubnetInfoBlock -JobName 'GetAdSubnetInfo' -VariableName "GetAdSubnetInfoResult" -ArgumentList $JobParameterArray | Out-Null

$JobParameterArray          = [Collections.ArrayList] @()
Init-PsJob -BlockName $CurrentNetworkBlock -JobName 'GetCurrentNetworkInfo' -VariableName "GetCurrentNetworkResult" -ArgumentList $JobParameterArray | Out-Null
#CurrentNetworkBlock


write-host "Fetching jobs...."

# All jobs are started, check if they are finished
While (get-job -HasMoreData $true)  
{
    foreach ($JobItem in $Global:JobArray)
    {
        if (($(Get-Variable -Name "$($JobItem.VarName)" -ValueOnly) -eq "") -and ((get-job -Name "$($JobItem.JobName)" -ChildJobState Running).State -eq 'Completed'))
        {
          
            switch ($JobItem.JobName) {
                "GetComputerInfo"       {   $Global:GetComputerInfoResult        = Receive-Job -Name "$($JobItem.JobName)"                                            
                                         #   $Global:GetComputerInfoResult       | Format-List                             
                                            write-host "--GetComputerInfoResult"        
                                        }
                "GetUserDomainInfo"     {   $Global:GetUserDomainInfoResult      = Receive-Job -Name "$($JobItem.JobName)"
                                         #   $Global:GetUserDomainInfoResult     | Format-List                                     
                                            write-host "--GetUserDomainInfoResult"
                                        }
                "GetMachineDomainInfo"  {   $Global:GetMachineDomainInfoResult   = Receive-Job -Name "$($JobItem.JobName)"
                                         #   $Global:GetMachineDomainInfoResult  | Format-List                                     
                                            write-host "--GetMachineDomainInfoResult"
                                        }
                "GetMachineWmiInfo"     {   $Global:GetMachineWmiInfoResult     = Receive-Job -Name "$($JobItem.JobName)"
                                         #   $Global:GetMachineWmiInfoResult     | Format-List                                     
                                            write-host "--GetMachineWmiInfoResult"
										}
				"GetAdSubnetInfo"   	{   $Global:GetAdSubnetInfoResult     = Receive-Job -Name "$($JobItem.JobName)"
										#   $Global:GetAdSubnetInfoResult     | Format-List                                     
										   write-host "--GetAdSubnetInfoResult"
                                           }
                "GetCurrentNetworkInfo" {   $Global:GetCurrentNetworkResult     = Receive-Job -Name "$($JobItem.JobName)"
										#   $Global:GetCurrentNetworkResult     | Format-List                                     
										   write-host "--GetCurrentNetworkResult"
									   	}
                Default {}
            }           
        }
    }
    Start-Sleep 1
}


