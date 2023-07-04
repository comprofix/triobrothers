<#
.SYNOPSIS
  Find Items that do not have an Image
.DESCRIPTION
  Find Items that do not have an Image path or missing image file
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021

.EXAMPLE
  .\Images_Report.ps1 -to <email_address>
  Specifiy Email Address to send the CSV File

#>

 if ($env:Processor_Architecture -ne "x86")
 { write-warning 'Launching x86 PowerShell'
 &"$env:WINDIR\syswow64\windowspowershell\v1.0\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line #-executionpolicy unrestricted
 exit
 }

 $MYDSN = "Micronet"
 $FROM_EMAIL = "support@triotrading.com.au"
 $TO_EMAIL = "emilyz@triotrading.com.au"
 $SMTP_SERVER = "mail.trio.local"


 function Connect-ODBC($mydsn,$query) {
   $dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"
   $conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
   $cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
   $da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
   $global:dt = New-Object system.Data.datatable
   $conn.Open() | Out-Null
   clear-host
   $global:null = $da.fill($dt)

 }


$query = "SELECT ITM_NO, ITM_DES, ITM_IMAGE FROM Inventory_Master_File WHERE ITM_STATUS = '0' and ITM_DES IS NOT NULL and ITM_DES <> '??' AND ITM_NO NOT LIKE '%PK%' AND ITM_NO NOT LIKE 'FS%' AND ITM_CAT NOT LIKE 'ESP%'"
Connect-ODBC $mydsn $query

$NO_IMAGES = $dt | Sort-Object ITM_NO

$FILE = "missing_images.csv"
$FileExists = Test-Path -Path $FILE
If ($FileExists -eq $True) {
  Remove-Item -Path $FILE -Force
}

#$NO_IMAGES

#Add-Content missing_images.csv "ITM_NO,ITM_DES,ITM_IMAGE"
$Output = @()

foreach($ITEM in $NO_IMAGES){
  $ITM_NO = $ITEM.ITM_NO
  $ITM_DES = $ITEM.ITM_DES
  $ITM_IMAGE = $ITEM.ITM_IMAGE

  if ([string]::IsNullOrEmpty($ITM_IMAGE)) {
  #Write-Host "No Image for $ITM_NO"
   $ITM_IMAGE = "NO IMAGE"
   $Output += New-Object -TypeName PSObject -Property @{
        ITM_NO = $ITM_NO
        ITM_IMAGE = $ITM_IMAGE
        ITM_DES = $ITM_DES
    } | Select-Object ITM_NO,ITM_DES,ITM_IMAGE
  }
  else
  {
	$FileExists = Test-Path -Path "$ITM_IMAGE" -ErrorAction SilentlyContinue
    if ($FileExists -eq $False) {
		#Write-Host "Image Missing"
      $ITM_IMAGE = $ITEM.ITM_IMAGE
	  $Output += New-Object -TypeName PSObject -Property @{
        ITM_NO = $ITM_NO
        ITM_IMAGE = $ITM_IMAGE
        ITM_DES = $ITM_DES
    } | Select-Object ITM_NO,ITM_DES,ITM_IMAGE
     }
  }

  #Store the information from this run into the array
}

$Output | Export-CSV -NoTypeInformation missing_images.csv

$Computer = $env:ComputerName

$Params = @{
   Subject = "Missing Images Report"
   Body = "See attached report for missing images"
   From = "$FROM_EMAIL"
   To = "$TO_EMAIL"
   SMTPSERVER = "$SMTP_SERVER"
   Attachment = "$FILE"
}


Send-MailMessage @Params
rm "$FILE"

exit
