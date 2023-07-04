$ErrorActionPreference = 'Stop'; # stop on all errors
$toolsDir   = "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
$fileLocation = Join-Path $toolsDir 'Micronet_ODBC_2.7_Setup.exe'
$filehelper = Join-Path $toolsDir 'Micronet_ODBC_2.7_Setup.iss'



$packageArgs = @{
  packageName   = $env:ChocolateyPackageName
  fileType      = 'EXE' #only one of these: exe, msi, msu
  file         = $fileLocation
  softwareName  = 'MicronetODBC*' #part or all of the Display Name as you see it in Programs and Features. It should be enough to be unique
  silentArgs   = "/S /f1`"$filehelper`""

}

# Install package
Install-ChocolateyInstallPackage @packageArgs # https://chocolatey.org/docs/helpers-install-chocolatey-install-package

# Create Registery Entries for odbc
if((Test-Path -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI") -ne $true) {  New-Item "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI" -force -ea SilentlyContinue  | out-null}
if((Test-Path -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\micronet") -ne $true) {  New-Item "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\micronet" -force -ea SilentlyContinue  | out-null}
if((Test-Path -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources") -ne $true) {  New-Item "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources" -force -ea SilentlyContinue  | out-null}
if((Test-Path -LiteralPath "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\TrioATrain") -ne $true) {  New-Item "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\TrioATrain" -force -ea SilentlyContinue  | out-null}
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\micronet' -Name 'Driver' -Value 'C:\Program Files (x86)\Micronet ODBC 2.7\bin\Micronet ODBC Client 2.7.dll' -PropertyType String -Force -ea SilentlyContinue | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\micronet' -Name 'Description' -Value 'odbc' -PropertyType String -Force -ea SilentlyContinue | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\micronet' -Name 'Database' -Value 'TrioA' -PropertyType String -Force -ea SilentlyContinue | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\micronet' -Name 'DescribeParam' -Value '' -PropertyType String -Force -ea SilentlyContinue | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources' -Name 'micronet' -Value 'Micronet Client Driver 2.7' -PropertyType String -Force -ea SilentlyContinue | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\ODBC Data Sources' -Name 'TrioATrain' -Value 'Micronet Client Driver 2.7' -PropertyType String -Force -ea SilentlyContinue | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\TrioATrain' -Name 'Driver' -Value 'C:\Program Files (x86)\Micronet ODBC 2.7\bin\Micronet ODBC Client 2.7.dll' -PropertyType String -Force -ea SilentlyContinue | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\TrioATrain' -Name 'Description' -Value 'TrioATrain' -PropertyType String -Force -ea SilentlyContinue  | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\TrioATrain' -Name 'Database' -Value 'TrioATrain' -PropertyType String -Force -ea SilentlyContinue  | out-null
New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI\TrioATrain' -Name 'DescribeParam' -Value '' -PropertyType String -Force -ea SilentlyContinue  | out-null

#Create Database INI File
Set-Content -Path "C:\Program Files (x86)\Micronet ODBC 2.7\schema\OADRD.INI" -value "[DB3]
ADDRESS=
PORT=
CONNECT_STRING=
TYPE=DB3
SCHEMA_PATH=
REMARKS=DB3
[TrioA]
ADDRESS=172.16.3.2
PORT=1706
CONNECT_STRING=,
TYPE=MSA
SCHEMA_PATH=
REMARKS=1
[TrioATrain]
ADDRESS=172.16.3.2
PORT=1706
CONNECT_STRING=,
TYPE=MSA
SCHEMA_PATH=
REMARKS=1
"
