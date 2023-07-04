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

  [Parameter(ParameterSetName = 'Contract', Mandatory = $true)]
  [Parameter(ParameterSetName = 'DSN', Mandatory = $true)]
  [String]$contract,

  [Parameter(ParameterSetName = 'AllContracts', Mandatory = $true)]
  [Parameter(ParameterSetName = 'DSN', Mandatory = $true)]
  [switch]$AllContracts,

  [Parameter(ParameterSetName = 'DSN', Mandatory = $true)]
  [string]$mydsn = "Micronet"
 )


 # Micronet ODBC Requires x86 Powershell, Run the script in x86 environment if run from x64 PowerShell
 if ($env:Processor_Architecture -ne "x86")
 {
   write-warning "Running x86 PowerShell..."
   &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noprofile $myinvocation.Line
   exit
 }


function DeleteContract {

  # Delete old Rolling Contracts
  Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Please Wait Deleting Old $CONTRACT Contract" -ForegroundColor Yellow

#Delete OBDC Query Statements
  $query = @(
    "DELETE FROM Contract_Line_File WHERE CONTL_NO LIKE '$CONTRACT%'",
    "DELETE FROM Contract_Application_File WHERE CONTA_NO LIKE '$CONTRACT%'",
    "DELETE FROM Contract_Header_File WHERE CONTH_NO LIKE '$CONTRACT%'"
  )

  foreach ($q in $query){
    $cmd = new-object System.Data.Odbc.OdbcCommand($q,$conn)
    $cmd.ExecuteNonQuery() | Out-Null
  }

}

function CreateContract {

  #Create new Rolling Contracts and assign debtor to each contract
  Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Please Wait while new $CONTRACT is Created." -ForegroundColor Yellow

  $query = @(
        "INSERT INTO Contract_Header_File (CONTH_NO, CONTH_DES, CONTH_TYPE) VALUES ('$CONTRACT','Contract for $description','0')"
          )

  foreach ($q in $query){
    $cmd = new-object System.Data.Odbc.OdbcCommand($q,$conn)
    $cmd.ExecuteNonQuery() | Out-Null
  }

}

function AddDebtorsContract {
  ForEach ($D in $DEBSEQ)  {
    $DBT_NO = $D.DBT_NO
    $CONTRACT = $D.CONTRACT
    $SEQNUM = $D.SEQNUM
    Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Adding Debtors $DBT_NO to $CONTRACT Contract" -ForegroundColor Yellow
    $q = "INSERT INTO Contract_Application_File (CONTA_NO, CONTA_DBTNO, CONTA_TYPE, CONTA_SEQ) VALUES ('$CONTRACT','$DBT_NO','0','$SEQNUM')"
    $cmd = new-object System.Data.Odbc.OdbcCommand($q,$conn)
    $cmd.ExecuteNonQuery() | Out-Null
  }

}

function AddItemsContract {
    ForEach ($I in $SEQ) {
      $ITM_NO = $I.ITM_NO
      $CONTRACT = $I.CONTRACT
      $ITM_SELL = $I.ITM_SELL
      $SEQNUM = $I.SEQNUM

      Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Adding item $ITM_NO to $CONTRACT Contract" -ForegroundColor Yellow

      $q = "INSERT INTO Contract_Line_File (CONTL_NO, CONTL_ITMNO, CONTL_TYPE, CONTL_DEFPRICE, CONTL_RETDEF, CONTL_TRADE0, CONTL_TRADE1,
            CONTL_TRADE2, CONTL_TRADE3, CONTL_TRADE4, CONTL_TRADE5, CONTL_TRADE6, CONTL_TRADE7, CONTL_RETAIL0, CONTL_RETAIL1,
            CONTL_RETAIL2, CONTL_RETAIL3, CONTL_RETAIL4, CONTL_RETAIL5, CONTL_RETAIL6, CONTL_RETAIL7, CONTL_SEQ) VALUES
            ('$CONTRACT','$ITM_NO','0','0','1','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL', '$ITM_SELL',
            '$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$SEQNUM')"

      $cmd = new-object System.Data.Odbc.OdbcCommand($q,$conn)
      $cmd.ExecuteNonQuery() | Out-Null
    }
}




write-host "Main script body"
clear-host

Remove-Item ".\tmp*.csv"

#"Always running in 32bit PowerShell at this point."
#$env:Processor_Architecture
#[IntPtr]::Size

# Delete old Rolling Contracts


#Set Variables
$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"
$categories_file = "$PSScriptRoot\contract_categories.csv"
$items_file = "$PSScriptRoot\contract_items.csv"
$debtors_file = "$PSScriptRoot\contract_debtors.csv"
$contracts_file = "$PSScriptRoot\contract_names.csv"


$items = $null
$products = $Null
$query = "SELECT ITM_NO, ITM_CAT, ITM_STATUS, ITM_ONHAND FROM Inventory_Master_File WHERE ITM_STATUS != 1"
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt = New-Object system.Data.datatable
$conn.Open() | Out-Null
Clear-Host
Write-Host $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") "Please Wait getting data" -ForegroundColor Yellow
$null = $da.fill($dt)
#$dt
$debtors = Import-CSV "$debtors_file"
$products = Import-CSV $categories_file
$items = Import-CSV $items_file
$contracts = Import-CSV $contracts_file



if ($PSBoundParameters.ContainsKey('delete')) {
  DeleteContract
  exit 0
}

if ($PSBoundParameters.ContainsKey('contract')) {
  $products = $products | Where-Object {$_.CONTRACT -eq $contract}
  $items = $items | Where-Object {$_.CONTRACT -eq $contract}
  $des = $Contracts | Where-Object {$_.CONTRACT -eq "$CONTRACT"}
  $description = $des.CON_DES
}

$cat = @()
foreach ($p in $products){

  $cat += $dt | Where-Object {$_.ITM_CAT -eq $p.ITM_CAT -and $_.ITM_ONHAND -gt 0} | Select-Object @{Name='CONTRACT';Expression={$p.CONTRACT}},ITM_NO,ITM_CAT,@{Name='ITM_SELL';Expression={$p.ITM_SELL}} #| Export-CSV "tmp_$contract.csv" -NoTypeInformation -Append

}

$itm = @()
foreach ($i in $items){

  $itm += $dt | Where-Object {$_.ITM_NO -eq $i.ITM_NO }  | Select-Object @{Name='CONTRACT';Expression={$i.CONTRACT}},ITM_NO,ITM_CAT,@{Name='ITM_SELL';Expression={$i.ITM_SELL}} #| Export-CSV "tmp_$contract.csv" -NoTypeInformation -Append

}


$itms = $itm + $cat


$SEQ = $itms | Group-Object -Property CONTRACT |
  ForEach-Object {
    $seqnum = 0
    foreach($item in $_.Group) {
        $seqnum += 10
        [pscustomobject]@{
            CONTRACT = $item.CONTRACT
            ITM_CAT = $item.ITM_CAT
            ITM_NO = $item.ITM_NO
            ITM_SELL  = $item.ITM_SELL
            SEQNUM =   $seqnum
        }
    }
}

$DEBSEQ = $debtors | Group-Object -Property CONTRACT |
  ForEach-Object {
    $seqnum = 0
    foreach($item in $_.Group) {
        $seqnum += 10
        [pscustomobject]@{
            CONTRACT = $item.CONTRACT
            DBT_NO = $item.DBT_NO
            SEQNUM =   $seqnum
        }
    }
}


if ($PSBoundParameters.ContainsKey('AllContracts')) {
  #$contracts = $debtors | Group-Object -Property CONTRACT |
  #  ForEach-Object {
  #    [pscustomobject]@{
  #        CONTRACT = $_.NAME
  #    }
  #  }

    foreach ($c in $Contracts) {
      $CONTRACT = $c.CONTRACT
      $des = $Contracts | Where-Object {$_.CONTRACT -eq "$CONTRACT"}
      $description = $des.CON_DES

      DeleteContract
      CreateContract
    }

    AddDebtorsContract
    AddItemsContract


} else {
  DeleteContract
  CreateContract
  $DEBSEQ = $DEBSEQ | Where contract -eq $contract
  AddDebtorsContract
  $SEQ = $SEQ | Where contract -eq $contract
  AddItemsContract
}


$conn.close()
Remove-Item ".\tmp*.csv"
