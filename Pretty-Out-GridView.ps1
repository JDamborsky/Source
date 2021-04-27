[string[]]$visible = 'Name', 'SID'
$type = 'DefaultDisplayPropertySet'
[Management.Automation.PSMemberInfo[]]$i =
New-Object System.Management.Automation.PSPropertySet($type,$visible)

Get-ADUser -LDAPFilter '(samaccountname=jar*)' |
    Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $i -Force -PassThru |
    Out-GridView -Title 'Select-User' -OutputMode Single |
    Select-Object -Property *