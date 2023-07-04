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
          USERID_NO as Staff_Code,
          USERID_NAME as First_Name,
          USERID_LASTNAME as Surname
          FROM Users_Master_File
          WHERE USERID_NO not LIKE 'PDAREP%'"

#$query = "Select * from Users_Master_File WHERE USERID_NO LIKE 'MATTHEW'"
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$conn.Open() | Out-Null
$dt += New-Object system.Data.datatable
$null += $da.fill($dt)

$query = "SELECT
          USERID_NAME as Staff_Code,
          USERID_NAME as First_Name,
          USERID_LASTNAME as Surname
          FROM Users_Master_File
          WHERE USERID_NO LIKE 'PDAREP%'"
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)



$null += $da.fill($dt)
$dt += New-Object system.Data.datatable
$dt | Export-CSV -NoTypeInformation -Path "$opmetrix_folder\Staff.csv"

$conn.close()
