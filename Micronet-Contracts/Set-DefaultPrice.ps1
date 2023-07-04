<#
.SYNOPSIS
  Set Pricebook Level
.DESCRIPTION
  Set Pricebook Level for all products
.PARAMETER pricelevel
  This is the number for the level you wish to set between 1 - 8
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  13/01/2023

.EXAMPLE
  .\Set-DefaultPrice.ps1 -pricelevel 1

#>


param (
  [Parameter(Mandatory = $true)]
  [int]$pricelevel,

  [string]$mydsn = "Micronet"
)



# Micronet ODBC Requires x86 Powershell, Run the script in x86 environment if run from x64 PowerShell
if ($env:Processor_Architecture -ne "x86") {
  write-warning "Running x86 PowerShell..."
  &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noprofile $myinvocation.Line
  exit
}

If ($pricelevel -eq '1') {
  $pricelevel = 0
} 
elseif ($pricelevel -eq '2') {
  $pricelevel = 1
}
elseif ($pricelevel -eq '3') {
  $pricelevel = 2
}
elseif ($pricelevel -eq '4') {
  $pricelevel = 3
}
elseif ($pricelevel -eq '5') {
  $pricelevel = 4
}
elseif ($pricelevel -eq '6') {
  $pricelevel = 5
}
elseif ($pricelevel -eq '7') {
  $pricelevel = 6
}
elseif ($pricelevel -eq '8') {
  $pricelevel = 7
}

$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"

#Get Debtors
$query = "SELECT DBT_NO FROM Debtors_Master_File"
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query, $conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt = New-Object system.Data.datatable
$conn.Open() | Out-Null
$null = $da.fill($dt)


$Debtors = $dt.DBT_NO

ForEach ($Debtor in $Debtors) {

  Write-Host "Updating Default Price Level for $Debtor" -ForegroundColor Yellow

  $query = "UPDATE Debtors_Master_File SET DBT_DEFPRICE = '$pricelevel' WHERE DBT_NO = '$Debtor'"
    
  $cmd = new-object System.Data.Odbc.OdbcCommand($query, $conn)
  $cmd.ExecuteNonQuery() | Out-Null
}

