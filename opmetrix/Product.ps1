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
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$conn.Open() | Out-Null
Clear-Host

$query = "SELECT
          ITM_NO as 'SKUCode',
          ITM_DES as 'Description',
          ITM_BC as 'Barcode',
          ITM_CLASS as 'Group',
          ITM_PRESUP as 'Vendor',
          ITM_TRADE0 as 'SellPrice1',
          ITM_TRADE1 as 'SellPrice2',
          ITM_TRADE2 as 'SellPrice3',
          ITM_TRADE3 as 'SellPrice4',
          ITM_TRADE4 as 'SellPrice5',
          ITM_TRADE5 as 'SellPrice6',
          ITM_TRADE6 as 'SellPrice7',
          ITM_TRADE7 as 'SellPrice8',
          ITM_STAX as 'TaxCode',
          '999' as 'TotalStockOnHand',
          ITM_SUNIT as 'DefaultSalesUnitQty',
          ITM_UNITS as 'DefaultSalesUnit',
          ITM_USER19 as 'Custom1',
          ITM_USER19 as 'Custom2',
          ITM_BUY0 as 'CostPricePerSalesUnit',
          ITM_CAT as 'ProductCategory1',
          CAT_DES as 'ProductCategory2',
          ITM_USER19 as 'ProductCategory3',
          ITM_USER19 as 'ProductCategory4',
          ITM_USER19 as 'ProductCategory5',
          ITM_USER19 as 'ProductCategory6',
          ITM_USER19 as 'ProductCategory7',
          ITM_USER19 as 'ProductCategory8',
          ITM_USER19 as 'PELCode',
          ITM_USER19 as 'FSACode',
          ITM_USER19 as 'FSWCode',
          ITM_USER19 as 'FSSICode',
          0 as 'HideInPortal'
          FROM Inventory_Master_File, Category_Description_Master_File
          WHERE ITM_STATUS = 0 AND
          (ITM_CAT = 'Z1' OR
          ITM_CAT = 'Z7' OR
          ITM_CAT LIKE '%Z') AND
          ITM_CAT = CAT_NO
          "
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt += New-Object system.Data.datatable
$null += $da.fill($dt)


$query = "SELECT
          ITM_NO as 'SKUCode',
          ITM_DES as 'Description',
          ITM_BC as 'Barcode',
          ITM_CLASS as 'Group',
          ITM_PRESUP as 'Vendor',
          ITM_TRADE0 as 'SellPrice1',
          ITM_TRADE1 as 'SellPrice2',
          ITM_TRADE2 as 'SellPrice3',
          ITM_TRADE3 as 'SellPrice4',
          ITM_TRADE4 as 'SellPrice5',
          ITM_TRADE5 as 'SellPrice6',
          ITM_TRADE6 as 'SellPrice7',
          ITM_TRADE7 as 'SellPrice8',
          ITM_STAX as 'TaxCode',
          ITM_ONHAND as 'TotalStockOnHand',
          ITM_SUNIT as 'DefaultSalesUnitQty',
          ITM_UNITS as 'DefaultSalesUnit',
          ITM_USER19 as 'Custom1',
          ITM_USER19 as 'Custom2',
          ITM_BUY0 as 'CostPricePerSalesUnit',
          ITM_CAT as 'ProductCategory1',
          CAT_DES as 'ProductCategory2',
          ITM_USER19 as 'ProductCategory3',
          ITM_USER19 as 'ProductCategory4',
          ITM_USER19 as 'ProductCategory5',
          ITM_USER19 as 'ProductCategory6',
          ITM_USER19 as 'ProductCategory7',
          ITM_USER19 as 'ProductCategory8',
          ITM_USER19 as 'PELCode',
          ITM_USER19 as 'FSACode',
          ITM_USER19 as 'FSWCode',
          ITM_USER19 as 'FSSICode',
          0 as 'HideInPortal'
          FROM Inventory_Master_File, Category_Description_Master_File
          WHERE ITM_STATUS = 0 AND
          ITM_CAT != 'Z2' AND
          ITM_CAT != 'A1' AND
          ITM_CAT != 'Z6' AND
          ITM_CAT != 'Z8' AND
          ITM_CAT != 'Z1' AND
          ITM_CAT != 'Z7' AND
          ITM_CAT != 'Z8A' AND
          ITM_CAT != 'Z9' AND
          ITM_CAT != 'ZEX' AND
          ITM_CAT != 'ZINT' AND
          ITM_CAT != 'CDSCT' AND
          (ITM_CAT NOT LIKE 'ESP%' AND
          ITM_CAT NOT LIKE '%Z') AND
          ITM_CAT = CAT_NO
          "
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$null += $da.fill($dt)
$dt += New-Object system.Data.datatable


$dt | Export-CSV -NoTypeInformation "$opmetrix_folder\Product.csv"

$conn.close()

$category = @(
  "I1",
  "I2",
  "I3",
  "I6",
  "J1",
  "J1K",
  "J1Z",
  "K1",
  "K2",
  "K3",
  "M1Z",
  "M5Z",
  "Z3",
  "Z3C1",
  "Z3C3",
  "Z3E1Q",
  "Z3F1",
  "Z3G2A",
  "Z3G3",
  "Z3H1",
  "Z3M1",
  "Z3V1A",
  "Z3V1B",
  "Z3V1C",
  "Z3V1D"
)

ForEach ($i in $category) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Clearance' | Set-Content "$opmetrix_folder\Product.csv"
}


$freight_subs = @(
  "Z7"
)

ForEach ($i in $freight_subs) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Freight Subs' | Set-Content "$opmetrix_folder\Product.csv"
}


$grinders = @(
  "G1"
)

ForEach ($i in $grinders) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Grinders' | Set-Content "$opmetrix_folder\Product.csv"
}


$hookahs = @(
  "E6",
  "E6A",
  "E7"

)

ForEach ($i in $hookahs) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Hookahs' | Set-Content "$opmetrix_folder\Product.csv"
}


$lighters = @(
  "C1",
  "C2",
  "C2A",
  "C2B",
  "C3",
  "C4",
  "C5",
  "C6",
  "C6A",
  "C7",
  "C8"
)

ForEach ($i in $lighters) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Lighters' | Set-Content "$opmetrix_folder\Product.csv"
}

$mobile = @(
  "M1"
)

ForEach ($i in $mobile) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Mobile Phone Accessories' | Set-Content "$opmetrix_folder\Product.csv"
}


$oil_pourers = @(
  "E8Q",
  "G10",
  "G9",
  "M2A"

)

ForEach ($i in $oil_pourers) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Oil Pourers' | Set-Content "$opmetrix_folder\Product.csv"
}

$readers = @(
  "A5",
  "A6",
  "A6A",
  "A5Z"

)

ForEach ($i in $readers) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Reading Glasses' | Set-Content "$opmetrix_folder\Product.csv"
}

$scales = @(
  "G5",
  "G7",
  "G7A",
  "G8",
  "H1",
  "H1A",
  "H2",
  "H3"

)

ForEach ($i in $scales) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Scales' | Set-Content "$opmetrix_folder\Product.csv"
}

$smokacc = @(
  "E1Q",
  "E1QA",
  "E5",
  "F1",
  "F1A",
  "F2",
  "F2A",
  "F3",
  "F3A",
  "F3B",
  "F3C",
  "F3D",
  "F3E",
  "F3F",
  "F3G",
  "F4",
  "G2",
  "G2A",
  "G3",
  "G4",
  "G6",
  "TMG6",
  "TP1"
)

ForEach ($i in $smokacc) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Smoking Accessories' | Set-Content "$opmetrix_folder\Product.csv"
}

$specials = @(
  "A12Z",
  "A6AZ",
  "A6Z",
  "C1Z",
  "C2Z",
  "C3Z",
  "C4Z",
  "C6Z",
  "C7Z",
  "C8Z",
  "E10Z",
  "E11Z",
  "E12Z",
  "E1QZ",
  "E1Z",
  "E4Z",
  "E6Z",
  "E8Z",
  "F1AZ",
  "F1Z",
  "F3BZ",
  "F3Z",
  "G1Z",
  "G2Z",
  "G3Z",
  "G5Z",
  "G6Z",
  "G7Z",
  "G9Z",
  "H1Z",
  "Z1",
  "Z1B"
)

ForEach ($i in $specials) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Specials' | Set-Content "$opmetrix_folder\Product.csv"
}


$stoneage = @(
  "E10"
)

ForEach ($i in $stoneage) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Stone Age' | Set-Content "$opmetrix_folder\Product.csv"
}

$sunglasses = @(
  "A10",
  "A10K",
  "A11",
  "A11K",
  "A11LK",
  "A11WI",
  "A12K",
  "A2K",
  "A3K",
  "A10Z",
  "A11LO",
  "A12",
  "A13GL",
  "A2",
  "A3",
  "A31",
  "A7",
  "B1",
  "B1A"
)

ForEach ($i in $sunglasses) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Sunglasses' | Set-Content "$opmetrix_folder\Product.csv"
}

$vapes = @(
  "TMV1A",
  "TMV1B",
  "TMV1C",
  "TMV1Z",
  "V1",
  "V1A",
  "V1AZ",
  "V1B",
  "V1BZ",
  "V1C",
  "V1CZ",
  "V1D",
  "V1DZ",
  "V1E",
  "V1EZ",
  "V1T",
  "V1Z",
  "V2",
  "V3",
  "V4",
  "V4A",
  "V4B",
  "V4C",
  "V4D",
  "V4E",
  "V4Z"
)

ForEach ($i in $vapes) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Vapes' | Set-Content "$opmetrix_folder\Product.csv"
}

$WP = @(
  "E12",
  "E1",
  "E11",
  "E12B",
  "E1A",
  "E2",
  "E2A",
  "E3",
  "E4",
  "E4A",
  "E8",
  "E9",
  "E9Z"
)

ForEach ($i in $WP) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'WP' | Set-Content "$opmetrix_folder\Product.csv"
}


$misc = @(
  "M5"
)

ForEach ($i in $misc) {
  (Get-Content "$opmetrix_folder\Product.csv") -Replace ('"'+ $i + '"'), 'Misc Products' | Set-Content "$opmetrix_folder\Product.csv"
}
"NOTE,Inactive Item,,,,,,,,,,,,,999,,,,,,,,,,,,,,,,,," | Add-Content "$opmetrix_folder\Product.csv"
