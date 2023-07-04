<#
.SYNOPSIS
  Sync Micronet to China
.DESCRIPTION
  This script will monitor the N:\Sync folder on china.triotrading.com.au for Copy of micronet Data. Extract the data and remove the zip file.DESCRIPTION
  Script should be setup as a service using NSSM. Install 7Zip on Server. Copy 7z.exe and 7z.dll to script location.
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/12/2021

#>

Write-Output "Started Running Script: $(Get-Date)"


$MICRONET_TEMP = "C:\Temp"
$remote_host = "nextdc.tbt.net.au"
$userName = 'trio'
$remote_password = get-content C:\scripts\china_password
$secureString = $remote_password | ConvertTo-SecureString -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $userName, $secureString
$hostkey = "SHA256:V80LOYaZ3IOpoUtGcwBT0ankPS1u8s0BCMebXaMlqkk"

Write-Output "Started Download from $remote_host $(Get-Date)"
pscp -l ${userName} -pw ${remote_password} -hostkey "$hostkey" ${userName}@${remote_host}:/home/trio/chinasync/chinasync.7z ${MICRONET_TEMP}

Write-Output "Download Completed: $(Get-Date)"
Write-Output "Start Extracting Files: $(Get-Date)"
7z x -i!Images $MICRONET_TEMP\chinasync.7z -aoa -oN:\Micronet
7z x -i!MNetLive $MICRONET_TEMP\chinasync.7z -aoa -oN:\MSA28
Write-Output "Extracting Completed: $(Get-Date)"
