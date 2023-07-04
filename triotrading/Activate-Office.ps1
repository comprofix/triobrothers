<#
.SYNOPSIS
  Active Office
.DESCRIPTION
  Active Office with KMS Server
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021
#>

clear-Host

$REG_PATH = "HKLM:\Software\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform"
$REG_NAME = "NoGenTicket"
$REG_VALUE = "1"
New-Item $REG_PATH -Force
New-ItemProperty -Path $REG_PATH -Name $REG_NAME -Value $REG_VALUE -PropertyType DWORD -Force

# Check if x86 or x64 for Office

$bitness = Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook -name Bitness

If($bitness -eq "x86") {\
  C:
  cd "\Program Files (x86)\Microsoft Office\Office16"
} else {
  #DO 64-BIT STUFF
  C:
  cd "\Program Files\Microsoft Office\Office16"
}

# Run Activation

$OFF_KEY = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
$KMS_SRV = "ca.tbt.net.au"

cscript //nologo ospp.vbs /inpkey:$OFF_KEY
cscript //nologo ospp.vbs /sethst:$KMS_SRV
cscript //nologo ospp.vbs /setprt:1688
cscript //nologo ospp.vbs /act
