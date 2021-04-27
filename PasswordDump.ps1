# important: this is required to load the assembly
[Windows.Security.Credentials.PasswordVault,Windows.Security.Credentials,ContentType=WindowsRuntime]


(New-Object Windows.Security.Credentials.PasswordVault).RetrieveAll() | 
ForEach-Object { 
    $_.RetrievePassword()
    $_ 
} | 
Select-Object -Property Username, Password, Resource |
Out-GridView 