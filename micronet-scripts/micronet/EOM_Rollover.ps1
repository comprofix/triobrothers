<#
.SYNOPSIS
  End of Month Roll Over
.DESCRIPTION
  This script will perform a backup of Micronet for End of Month Rollover
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021

#>

[CmdletBinding()]
param (
  [switch] $global:verbose
 )

 $ACCOUNTS_EMAIL="minc@triotrading.com.au"
 $COMPUTERNAME = $env:computername
 $MICRONET_TEMP="N:\TEMP"
 $ARCHIVE_FOLDER = "\\archive\archive\Micronet\Archives"
 $current_year = $(Get-Date -Format yyyy)
 $previous_month = $((get-date).date.AddMonths(-1).ToString("MM"))
 $date_format = $(Get-Date -Format yyyyMMdd)

 If ($date_format -eq ($current_year + "0101")) {
    $current_year = ($current_year - 1)
 }


 $MICRONETSERVICES = @("MicronetDConnect28NZ",
               "MicronetDConnect28NZintoAUS",
               "MicronetDConnect28 NZ Repl",
               "MicronetDConnect28 - 2",
               "MicronetDConnect28",
               "MicronetLLS27",
               "MicronetLLS28",
               "MSASS12",
               "MSASS11"
             )



 $global:VerbosePreference = 'SilentlyContinue'
 if($verbose) {
    $VerbosePreference = "continue"
   }

 #write-host $COMPUTERNAME

 If ($COMPUTERNAME -eq "MICRONET") {
   Write-Verbose "Connected to Micronet. Continuing."
 } else {
	 Write-host "NOT Connected to Micronet. Exiting."
	 exit 0
 }

clear-Host

function RemoveCopyBackupFolders {
	Write-Verbose "Removing old N:\Temp backup files"
	Remove-Item $MICRONET_TEMP\masters -Recurse -Force -ErrorAction SilentlyContinue
	Remove-Item $MICRONET_TEMP\MnetLive -Recurse -Force -ErrorAction SilentlyContinue
}


function CopyFolders {
	Write-Verbose "Please Wait. Copying files to N:\Temp\"
	Copy-Item -Path N:\MSA28\masters -Destination $MICRONET_TEMP\masters -Recurse
	Copy-Item -Path N:\MSA28\MnetLive -Destination $MICRONET_TEMP\MNetLive -Recurse
}

function ArchiveFolder {

  If (Test-Path $backup_path -PathType Any) {
    }
  Else {
    mkdir "$archive_folder\$current_year\$previous_month-$current_year"
  }

  Set-Location N:\TEMP\
  if($VerbosePreference -eq "continue" ) {
    & N:\MSA28\7z.exe a -r $MICRONET_TEMP\masters-$date_format.7z $MICRONET_TEMP\Masters
    & N:\MSA28\7z.exe a -r $MICRONET_TEMP\MnetLive-$date_format.7z $MICRONET_TEMP\MnetLive
  }
  Else {
    & N:\MSA28\7z.exe -bso0 -bsp0 a -r $MICRONET_TEMP\masters-$date_format.7z $MICRONET_TEMP\Masters
    & N:\MSA28\7z.exe -bso0 -bsp0 a -r $MICRONET_TEMP\MnetLive-$date_format.7z $MICRONET_TEMP\MnetLive
  }

  Copy-Item $MICRONET_TEMP\masters-$date_format.7z $backup_path
  Copy-Item $MICRONET_TEMP\MnetLive-$date_format.7z $backup_path

}


$backup_path = "$archive_folder\$current_year\$previous_month-$current_year"

$Params = @{
  Subject = "EOM Micronet Backup Completed"
  Body = "EOM Micronet Backup Completed"
  To = "$ACCOUNTS_EMAIL"
  FROM = "support@triotrading.com.au"
  SMTPSERVER = "mail.trio.local"
}



Write-Verbose "Stopping Services..."
If ($computername -eq "micronet") {
  ForEach ($Service in $MICRONETSERVICES){
    Stop-Service "$Service"
  }
  RemoveCopyBackupFolders
  CopyFolders
  Write-Verbose "Starting Services..."
  ForEach ($Service in $MICRONETSERVICES){
    Start-Service "$Service"
  }
  Send-MailMessage @Params
  ArchiveFolder
  RemoveCopyBackupFolders

}
Else {
  exit
}

Write-Verbose "Miconet Backup Completed"
