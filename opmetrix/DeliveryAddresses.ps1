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
$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"
$items = $null
$products = $Null
$opmetrix_folder = "N:\Opmetrix\Import"


$query = "SELECT
          DEL_DBTNO as CustomerCode,
          DEL_DELNO as CustomersDeliveryCode,
          DEL_DELADR0 as DeliveryAddressLine1,
          DEL_DELADR1 as DeliveryAddressLine2,
          DEL_DELADR2 as DeliveryAddressLine3
          FROM Debtors_Delivery_Address_File, Debtors_Master_File
          WHERE DEL_DBTNO = DBT_NO AND DBT_STATUS=0"

#$query = "Select * from Users_Master_File WHERE USERID_NO LIKE 'MATTHEW'"
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$conn.Open() | Out-Null
$dt += New-Object system.Data.datatable
$null += $da.fill($dt)

$dt += New-Object system.Data.datatable
$dt | Export-CSV -NoTypeInformation -Path "$opmetrix_folder\DeliveryAddresses.csv"

$conn.close()
