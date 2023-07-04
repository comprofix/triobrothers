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
          DBT_NO as 'CustomerID',
          DBT_NAME as 'CustomerName',
          (DBT_DEFPRICE + 1) as 'PriceLevel',
          DBT_DSCNT as 'CustomerDiscount',
          DBT_POSTADR0 as 'AddressLine1',
          DBT_POSTADR1 as 'AddressLine2',
          DBT_POSTADR2 as 'AddressLine3',
          DBT_POSTADR3 as 'AddressLine4',
          DBT_POSTPC as 'PostCode',
          DBT_FAX as 'Fax',
          DBT_PHONE as 'Phone',
          DBT_USER19 as 'Notes',
          DBT_SALESMAN as 'StaffCode',
          DBT_DSCNT as 'SpecialPricingGroup',
          DBT_ACCHOLD as 'StopCredit',
          '' as 'CashOnlyFlag',
          DBT_MOBILE as 'Mobile',
          DBT_INTERNET as 'Email',
          DBT_CURBAL3 as 'Balance3',
          DBT_CURBAL2 as 'Balance2',
          DBT_CURBAL1 as 'Balance1',
          DBT_CURBAL0 as 'BalanceCurrent',
          DBT_CURRENT as 'BalanceTotal',
          0 as 'CompulsoryOrderNo',
          1 as 'PrintPricing',
          DBT_USER19 as 'Vendor',
          1 as 'EnableEditPrice',
          DBT_CLASS as 'CustomerCategory1',
          DBT_STER as 'CustomerCategory2',
          DBT_USER19 as 'CustomerCategory3',
          DBT_USER19 as 'CustomerCategory4',
          DBT_USER19 as 'CustomerCategory5',
          DBT_USER19 as 'CustomerCategory6',
          DBT_USER19 as 'CustomerCategory7',
          DBT_USER19 as 'CustomerCategory8',
          DBT_USER19 as 'CustomField1',
          DBT_USER19 as 'CustomField2',
          '' as 'ReportOnly'
          FROM Debtors_Master_File
          WHERE DBT_STATUS=0 AND DBT_NAME NOT LIKE 'USE%'"


$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt = New-Object system.Data.datatable
$conn.Open() | Out-Null
Clear-Host
Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Please Wait getting data" -ForegroundColor Yellow
$null = $da.fill($dt)
$dt | Export-CSV -NoTypeInformation "$opmetrix_folder\Customer.csv"
$conn.close()
