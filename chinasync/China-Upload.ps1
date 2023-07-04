<#
.SYNOPSIS
  Sync Micronet to China
.DESCRIPTION
  This script will perform a backup of Micronet and copy to china.triotrading.com.au
.NOTES
  Author:         Matthew McKinnon
  Creation Date:  01/12/2021

#>

[CmdletBinding()]
param (
  [switch] $global:verbose
 )

clear-Host

$MICRONET_TEMP = "N:\chinasync"
$backup_path = "$MICRONET_TEMP\mnetlive"
$images_path = "$MICRONET_TEMP\images"
$remote_host = "nextdc.tbt.net.au"
$china_host = "china.triotrading.com.au"
$userName = 'trio'
$remote_password = get-content C:\scripts\china_password
$secureString = $remote_password | ConvertTo-SecureString -AsPlainText -Force
$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $userName, $secureString
$hostkey = "SHA256:V80LOYaZ3IOpoUtGcwBT0ankPS1u8s0BCMebXaMlqkk"

Write-Output "Started Running Script: $(Get-Date)"

$global:VerbosePreference = 'SilentlyContinue'
if($verbose) {
  $VerbosePreference = "Continue"
}

If (!(Test-Path $MICRONET_TEMP -PathType Any)) {
  Write-Verbose "$MICRONET_TEMP does not exist. Creating Folder"
  New-Item $MICRONET_TEMP -Type Directory | Out-Null
}

Write-Output "Please Wait. Copying files to $MICRONET_TEMP $(Get-Date) "
robocopy N:\Micronet\Images $MICRONET_TEMP\Images /MIR
robocopy N:\MSA28\MnetLive $MICRONET_TEMP\MnetLive /MIR
Write-Output "Archiving Files: $(Get-Date) "
7z a $MICRONET_TEMP\chinasync.7z $backup_path $images_path
Write-Output "Uploading Files: $(Get-Date) "
pscp -l ${userName} -pw ${remote_password} -hostkey ${hostkey} ${MICRONET_TEMP}\chinasync.7z* ${userName}@${remote_host}:/home/trio/chinasync
Write-Output "Starting Remote Download: $(Get-Date) "
Write-Output "Completed: $(Get-Date)"
