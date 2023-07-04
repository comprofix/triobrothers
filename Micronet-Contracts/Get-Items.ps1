<#
.SYNOPSIS
  Create Customer Contracts in Micronet
.DESCRIPTION
  Contract_folders is a script that will import the customer contracts that have been configured in each folder.
.PARAMETER contract
  This is the folder name for the contract.
.PARAMETER centdebt
  This is the Central Debtor name
.PARAMETER mydsn
  Use this to change the ODBC DSN when you want to test the script in Training
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  17/06/2020

.EXAMPLE
  .\Contract_folders.ps1 -contract BARCITY
  Create Contract for BARCITY

.EXAMPLE
  .\Contract_folders.ps1 -contract BARCITY -mydsn TrioATrain
  Create Contract for BARCITY into the Training Database

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



write-host "Main script body"
clear-host


#Set Variables
$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"
$query = "SELECT ITM_NO, ITM_CAT, ITM_STATUS FROM Inventory_Master_File WHERE ITM_STATUS=0"
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt = New-Object system.Data.datatable
$conn.Open() | Out-Null
Clear-Host
Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Please Wait getting data" -ForegroundColor Yellow
$null = $da.fill($dt)
$dt




$conn.close()
Remove-Item ".\tmp*.csv"
