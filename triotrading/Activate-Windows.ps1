<#
.SYNOPSIS
  Active Windows
.DESCRIPTION
  Active Windows with KMS Server
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

CD \WINDOWS\SYSTEM32

$WIN_KEY = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
$KMS_SRV = "ca.tbt.net.au:1688"

cscript //nologo slmgr.vbs /upk
cscript //nologo slmgr.vbs /ipk $WIN_KEY
cscript //nologo slmgr.vbs /skms $KMS_SRV
cscript //nologo slmgr.vbs /ato
