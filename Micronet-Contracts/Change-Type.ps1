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
  [switch]$delete,

  [Parameter(ParameterSetName = 'DSN', Mandatory = $false)]
  [string]$mydsn = "Micronet"
)


# Micronet ODBC Requires x86 Powershell, Run the script in x86 environment if run from x64 PowerShell
if ($env:Processor_Architecture -ne "x86") {
  write-warning "Running x86 PowerShell..."
  &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noprofile $myinvocation.Line
  exit
}

#Set Variables
$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query, $conn)
$conn.Open() | Out-Null


# Delete old Rolling Contracts
Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Please Wait Deleting Old $CONTRACT Contract" -ForegroundColor Yellow

$query = "DELETE FROM Contract_Header_File WHERE CONTH_NO LIKE '%'"


  $cmd = new-object System.Data.Odbc.OdbcCommand($query, $conn)
  $cmd.ExecuteNonQuery() | Out-Null

