$PathConfig = 'C:\zabbix\zabbix_agentd.conf'
$Name = 'agent-name'
$HostName = 'Hostname=' + $Name
$Server = 'zabbix-server'
$ServerActive = 'ServerActive=' + $Server

Stop-Service -Name "zabbix-agent" -ErrorAction SilentlyContinue
sc.exe delete "zabbix-agent"

if (Test-Path -Path 'c:\zabbix') {
            Write-Host Directory not empty
            Rename-Item c:\zabbix c:\zabbix.old
    }
New-Item -ItemType Directory -Path c:\zabbix
Invoke-WebRequest -Uri "https://cdn.zabbix.com/zabbix/binaries/stable/5.2/5.2.5/zabbix_agent-5.2.5-windows-amd64-openssl.zip" -OutFile "C:\zabbix\zabbix_agent-5.2.5-windows-amd64-openssl.zip"

[System.Reflection.Assembly]::LoadWithPartialName('System.IO.Compression.FileSystem')
[System.IO.Compression.ZipFile]::ExtractToDirectory('C:\zabbix\zabbix_agent-5.2.5-windows-amd64-openssl.zip', 'c:\zabbix')

$config = ('Server=', 'LogFile=C:\zabbix\zabbix_agentd.log', 'LogFileSize=0', 'ServerActive=baa-group.net', 'StartAgents=0')
$config | Out-File $PathConfig
Add-Content $PathConfig $HostName
Add-Content $PathConfig $ServerActive

$MyRawString = Get-Content -Raw $PathConfig
$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
[System.IO.File]::WriteAllLines($PathConfig, $MyRawString, $Utf8NoBomEncoding)

New-Service -Name "zabbix-agent" -BinaryPathName 'c:\zabbix\bin\zabbix_agentd.exe --config "c:\zabbix\zabbix_agentd.conf"' -StartupType Automatic
Start-Service -Name "zabbix-agent"
