Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
choco source add -n="Trio" -s="http://choco.tbt.net.au/nuget/trio"
choco source add -n="Community" -s="http://choco.tbt.net.au/nuget/community" --priority=10
choco source disable -n="Chocolatey"
choco install boxstarter -y
Import-Module C:\ProgramData\Boxstarter\BoxStarterShell.ps1
Set-BoxstarterConfig -NugetSources "http://choco.tbt.net.au/nuget/trio;http://choco.tbt.net.au/nuget/community" | Out-Null
Install-BoxStarterPackage -PackageName http://git.tbt.net.au/mmckinnon/chocolatey-packages/-/raw/master/boxstarter.ps1 -Credential TRIO\matthew_sa
