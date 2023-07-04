<#
.SYNOPSIS
  Opmetrix Priving Export
.DESCRIPTION
  Export Data for Opmetrix Schema and Import
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  25/05/2022
#>

param (
  [string]$mydsn = "Micronet"
 )


 # Micronet ODBC Requires x86 Powershell, Run the script in x86 environment if run from x64 PowerShell
 if ($env:Processor_Architecture -ne "x86")
 {
   write-warning "Running x86 PowerShell..."
   &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noprofile $myinvocation.Line
   exit
 }

clear-host

#Set Variables
$todate = Get-Date -Format "yyyyMMdd"
$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"
$items = $null
$products = $Null
$opmetrix_orders = "N:\Opmetrix\Export"
$Processed = "N:\Opmetrix\Processed"
$micronet_orders = "N:\MSA28\MnetLive\A\pda\Orders\opmetrix"
$newline = @()


$files = Get-ChildItem -Path $opmetrix_orders -Name
foreach ($file in $files) {
  $content = get-content $opmetrix_orders\$file
  foreach ($line in $content) {
    $array = $line.split(",")
    If ($array[0] -eq "H") {

      $custcode = $array[6]
      $adrcode = $array[15]
      $rep = $array[11]
      $on = $array[1]
      $array[10] = $rep + '%' + $on
      $fname = $todate + "_" + $array[10]

      $query = "SELECT
                DEL_DBTNO,
                DEL_DELNO,
                DEL_DELADR0,
                DEL_DELADR1,
                DEL_DELADR2,
                DEL_DELADR3,
                DEL_DELADR4,
                DEL_POSTCODE,
                DEL_FRGTNO
                FROM Debtors_Delivery_Address_File
                WHERE DEL_DBTNO='$CUSTCODE' and
                DEL_DELNO = '$adrcode' "

      #$query = "Select * from Users_Master_File WHERE USERID_NO LIKE 'MATTHEW'"
      $conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
      $cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
      $da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
      $conn.Open() | Out-Null
      $dt = New-Object system.Data.datatable
      $null = $da.fill($dt)

      $firstline = $array -join ","

      $firstline + ',"' + $dt.DEL_DELADR0 + '","' + $dt.DEL_DELADR1 + '","' + $dt.DEL_DELADR2 + '","' + $dt.DEL_POSTCODE + '","' + $dt.DEL_DELADR4 + '","' + $dt.DEL_FRGTNO + '"' | Set-Content $micronet_orders\"$fname".txt


    }
    elseif ($array[0] -eq "D" -and $array[2] -ne "NOTE") {
      $line | Add-Content $micronet_orders\"$fname".txt
      If ($array[10] -ne "") {
        'T,' + $array[1] + ',' + $array[10] | Add-Content $micronet_orders\"$fname".txt
      }
    } elseif ($array[2] -eq "NOTE" ) {
      $newline += 'T,' + $array[1] + ',' + $array[10]
    }
  }
  $newline | Add-Content $micronet_orders\"$fname".txt
  Move-Item -Path $opmetrix_orders\$file -Destination $Processed

}
