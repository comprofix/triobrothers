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
$dte = Get-Date
$dte = $dte.AddDays(-5)
$DATE = Get-Date $dte -Format yyyy/MM/dd
$opmetrix_folder = "N:\Opmetrix\Import"

$query = "SELECT
          DIH_INVNO as 'ExternalTransactionID',
          '' as 'OpmetrixID',
          'C' as 'TransactionType',
          DIH_DATE as 'DeliveryDate',
          DIH_DBTNO as 'CustomerCode',
          DBT_NAME as 'CustomerName',
          DIH_WARENO as 'StockLocation',
          DIH_INVNO as 'OurReference',
          DIH_ORDNO as 'OrderNumber',
          DIH_USERID as 'StaffCode',
          ((DIH_INVAMT*-1)-(DIH_INVAMT*-1/11)) as 'InvoiceTotal',
          (DIH_INVAMT*-1/11) as 'TaxTotal',
          DIH_DELNO as 'DeliveryAddressCode',
          DIH_DATE as 'CapturedDate',
          DIL_ITMNO as 'ProductCode',
          DIL_DES as 'ProductDescription',
          DIL_QTYDEL as 'QuantitySold',
          DIL_SELL as 'UnitPrice',
          DIL_DSCNT as 'DicountPercent',
          '' as 'DiscountAmount',
          DIL_ORDERTOTAL as 'LineTotal',
          (DIL_SELLTOTAL/11) as 'LineTax',
          '' as 'Note'
          FROM Debtors_Invoice_Header_File, Debtors_Invoice_Line_File, Debtors_Master_File
          WHERE
          DIL_DIHLINK = DIH_LINK AND
          DIH_DATE > '$date' AND
          DIH_TYPE = '67' AND
          DBT_NO = DIH_DBTNO"

#$query = "SELECT * FROM Warehouse_Master_File"

$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt += New-Object system.Data.datatable
$conn.Open() | Out-Null
Clear-Host
$null += $da.fill($dt)

$query = "SELECT
          DIH_INVNO as 'ExternalTransactionID',
          '' as 'OpmetrixID',
          'I' as 'TransactionType',
          DIH_DATE as 'DeliveryDate',
          DIH_DBTNO as 'CustomerCode',
          DBT_NAME as 'CustomerName',
          DIH_WARENO as 'StockLocation',
          DIH_INVNO as 'OurReference',
          DIH_ORDNO as 'OrderNumber',
          DIH_USERID as 'StaffCode',
          (DIH_INVAMT-(DIH_INVAMT/11)) as 'InvoiceTotal',
          (DIH_INVAMT/11) as 'TaxTotal',
          DIH_DELNO as 'DeliveryAddressCode',
          DIH_DATE as 'CapturedDate',
          DIL_ITMNO as 'ProductCode',
          DIL_DES as 'ProductDescription',
          DIL_QTYDEL as 'QuantitySold',
          DIL_SELL as 'UnitPrice',
          DIL_DSCNT as 'DicountPercent',
          '' as 'DiscountAmount',
          DIL_ORDERTOTAL as 'LineTotal',
          (DIL_SELLTOTAL/11) as 'LineTax',
          '' as 'Note'
          FROM Debtors_Invoice_Header_File, Debtors_Invoice_Line_File, Debtors_Master_File
          WHERE
          DIL_DIHLINK = DIH_LINK AND
          DIH_DATE > '$date' AND
          DIH_TYPE = '73' AND
          DBT_NO = DIH_DBTNO"

$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$null += $da.fill($dt)
$dt += New-Object system.Data.datatable
#$dt | Export-CSV -NoTypeInformation -Path "$opmetrix_folder\Transactions.csv"
($dt | ConvertTo-Csv -NoTypeInformation) | Select-Object -Skip 1 | Set-Content -Path "$opmetrix_folder\Transactions.csv"

$conn.close()
