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
          COMP_TAX0,
          COMP_TAX1
          FROM Company_Control_File
          WHERE COMP_NO = 'A'"


$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt = New-Object system.Data.datatable
$conn.Open() | Out-Null
Clear-Host
$null = $da.fill($dt)

$taxrates = @(
    [PSCustomObject]@{
        Tax_Code = "1"
        Description = "GST"
        Rate = $dt.COMP_TAX0
    },

    [PSCustomObject]@{
      Tax_Code = "2"
      Description = "GSTFREE"
      Rate = $dt.COMP_TAX1
    }
)

$taxrates | Export-CSV -NoTypeInformation -Path "$opmetrix_folder\Tax.csv"

$conn.close()
