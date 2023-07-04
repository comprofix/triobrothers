<#
.SYNOPSIS
  Extract PODS from XL Express
.DESCRIPTION
  Extract ZIP File of Proof of Delivery Sheets (PODS) provided by XL Express
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021
#>

$ZIP = Get-ChildItem -Path \\trio.local\Data\Archive\PODS\*.zip -Name

Foreach ($File in $ZIP) {
  Expand-Archive -LiteralPath \\trio.local\Data\Archive\PODS\$FILE -DestinationPath \\trio.local\data\Shared\PODS -Force
  Move-Item -Path \\trio.local\Data\Archive\PODS\$FILE -Destination \\trio.local\Data\Archive\PODS\Archived\$FILE
}
