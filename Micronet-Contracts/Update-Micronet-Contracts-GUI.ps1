<# This form was created using POSHGUI.com  a free online gui designer for PowerShell
.NAME
    Untitled
#>

$mydsn = "Micronet"

# Micronet ODBC Requires x86 Powershell, Run the script in x86 environment if run from x64 PowerShell
if ($env:Processor_Architecture -ne "x86")
{
  write-warning "Running x86 PowerShell..."
  &"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" -noprofile $myinvocation.Line
  exit
}

function new_contract {

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  [System.Windows.Forms.Application]::EnableVisualStyles()

  $Form                            = New-Object system.Windows.Forms.Form
  $Form.ClientSize                 = New-Object System.Drawing.Point(400,100)
  $Form.text                       = "Create New Micronet Contract"
  $Form.TopMost                    = $true


  $lblContract                     = New-Object system.Windows.Forms.Label
  $lblContract.text                = "Enter Contract Name"
  $lblContract.AutoSize            = $true
  $lblContract.width               = 30
  $lblContract.height              = 10
  $lblContract.location            = New-Object System.Drawing.Point(10,12)
  $lblContract.Font                = New-Object System.Drawing.Font('Verdana',10)
  $Form.Controls.Add($lblContract)

  $TextBox1                        = New-Object system.Windows.Forms.TextBox
  $TextBox1.multiline              = $false
  $TextBox1.width                  = 100
  $TextBox1.height                 = 20
  $TextBox1.location               = New-Object System.Drawing.Point(175,12)
  $TextBox1.Font                   = New-Object System.Drawing.Font('Verdana',10)

  $Form.controls.AddRange(@($TextBox1))



  $OKButton                         = New-Object system.Windows.Forms.Button
  $OKButton.text                    = "OK"
  $OKButton.width                   = 50
  $OKButton.height                  = 30
  $OKButton.location                = New-Object System.Drawing.Point(10,50)
  $OKButton.Font                    = New-Object System.Drawing.Font('Verdana',10)
  $Form.Controls.Add($OKButton)

  $OKButton.Add_Click({
    $Global:Contract = $TextBox1.Text
    mkdir $Contract
    Copy-Item -Path $PSScriptRoot\TEMPLATE\*.csv $PSScriptRoot\$Contract
    $form.Close()
  })

  [void] $Form.ShowDialog()

}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function Append-ColoredLine {
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [System.Windows.Forms.RichTextBox]$box,
        [Parameter(Mandatory = $true, Position = 1)]
        [System.Drawing.Color]$color,
        [Parameter(Mandatory = $true, Position = 2)]
        [string]$text
    )
    $box.SelectionStart = $box.TextLength
    $box.SelectionLength = 0
    $box.SelectionColor = $color
    $box.AppendText($text)
    $box.AppendText([Environment]::NewLine)
    $box.ScrollToCaret()
}

$dsn = "DSN=$mydsn;UID=odbc;PWD=odbc;"
$items = $null
$products = $Null
$query = "SELECT ITM_NO, ITM_CAT FROM Inventory_Master_File"
$conn = New-Object System.Data.Odbc.OdbcConnection($dsn)
$cmd = New-object System.Data.Odbc.OdbcCommand($query,$conn)
$da = New-Object system.Data.odbc.odbcDataAdapter($cmd)
$dt = New-Object system.Data.datatable
$conn.Open() | Out-Null
Clear-Host
$null = $da.fill($dt)



$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(1000,600)
$Form.text                       = "Update Micronet Contracts"
$Form.TopMost                    = $false

$richText = New-Object System.Windows.Forms.RichTextBox
$richText.Location = [System.Drawing.Point]::new(10,250)
$richText.Size = [System.Drawing.Size]::new(950,300)
$richText.BackColor = [System.Drawing.Color]::FromArgb(0,0,0)
$richText.ForeColor = [System.Drawing.Color]::FromArgb(255,255,255)
$richText.Font = [System.Drawing.Font]::new('Lucida Console',14)
$richText.Anchor = 'Top','Right','Bottom','Left'
$form.Controls.Add($richText)

$lblContract                     = New-Object system.Windows.Forms.Label
$lblContract.text                = "Contract"
$lblContract.AutoSize            = $true
$lblContract.width               = 30
$lblContract.height              = 10
$lblContract.location            = New-Object System.Drawing.Point(10,12)
$lblContract.Font                = New-Object System.Drawing.Font('Verdana',10)
$Form.Controls.Add($lblContract)

$lstContract                       = New-Object system.Windows.Forms.ComboBox
$lstContract.text                  = ""
$lstContract.width                 = 120
$lstContract.height                = 20
Get-ChildItem -Path $PSScriptRoot -Directory | Select Name | ForEach-Object {[void] $lstContract.Items.Add($_.Name)}
$lstContract.location              = New-Object System.Drawing.Point(80,12)
$lstContract.Font                  = New-Object System.Drawing.Font('Verdana',10)
$Form.Controls.Add($lstContract)

$grpbox                            = New-Object System.Windows.Forms.GroupBox
$grpbox.location                   =  New-Object System.Drawing.Point(10,50)
$grpbox.width                = 120
$grpbox.height               = 100
$form.controls.add($grpbox)

$rb                     = New-Object system.Windows.Forms.RadioButton
$rb.text                 = "debtors"
$rb.AutoSize             = $true
$rb.width                = 104
$rb.height               = 20
$rb.location             = New-Object System.Drawing.Point(10,10)
$rb.Font                 = New-Object System.Drawing.Font('Verdana',10)
$grpbox.controls.add($rb)

$rb                     = New-Object system.Windows.Forms.RadioButton
$rb.text                 = "items"
$rb.AutoSize             = $true
$rb.width                = 104
$rb.height               = 20
$rb.location             = New-Object System.Drawing.Point(10,30)
$rb.Font                 = New-Object System.Drawing.Font('Verdana',10)
$grpbox.controls.add($rb)

$rb                     = New-Object system.Windows.Forms.RadioButton
$rb.text                 = "categories"
$rb.AutoSize             = $true
$rb.width                = 104
$rb.height               = 20
$rb.location             = New-Object System.Drawing.Point(10,50)
$rb.Font                 = New-Object System.Drawing.Font('Verdana',10)
$grpbox.controls.add($rb)


$ViewButton                         = New-Object system.Windows.Forms.Button
$ViewButton.text                    = "View Contract"
$ViewButton.width                   = 200
$ViewButton.height                  = 30
$ViewButton.location                = New-Object System.Drawing.Point(220,10)
$ViewButton.Font                    = New-Object System.Drawing.Font('Verdana',10)
$Form.Controls.Add($ViewButton)

$ViewButton.Add_Click({
  $richText.Clear()
  $Contract = $lstContract.SelectedItem
  $r = $grpbox.Controls | Where-Object{ $_.Checked } | Select-Object Text
  $t = $r.text
  If (!$Contract -or !$r) {
    Append-ColoredLine $richText Red "Cotract or details not selected."
  } else
  {
    $data = Import-CSV ".\$Contract\$t.csv"
    If (!$data) {
      Append-ColoredLine $richText Red "Contract has no $t"
    } else
    {
      $data | Out-GridView -Title "$t on $Contract Contract"
    }
  }

})

$EditButton                         = New-Object system.Windows.Forms.Button
$EditButton.text                    = "Edit Contract"
$EditButton.width                   = 200
$EditButton.height                  = 30
$EditButton.location                = New-Object System.Drawing.Point(220,50)
$EditButton.Font                    = New-Object System.Drawing.Font('Verdana',10)
$Form.Controls.Add($EditButton)

$EditButton.Add_Click({
  $richText.Clear()
  $Contract = $lstContract.SelectedItem
  $r = $grpbox.Controls | Where-Object{ $_.Checked } | Select-Object Text
  $t = $r.text
  If (!$Contract -or !$r) {
    Append-ColoredLine $richText Red "Cotract or details not selected."
  } else
  {
    Invoke-Item "$PSScriptRoot\$Contract\$t.csv"
  }
})

$NewButton                        = New-Object system.Windows.Forms.Button
$NewButton.text                    = "New Contract"
$NewButton.width                   = 200
$NewButton.height                  = 30
$NewButton.location                = New-Object System.Drawing.Point(220,130)
$NewButton.Font                    = New-Object System.Drawing.Font('Verdana',10)
$Form.Controls.Add($NewButton)

$NewButton.Add_Click({
  new_contract
  $lstContract.Items.Clear()
  $lstContract.BeginUpdate()
  Append-ColoredLine $richText Green "New contract files for $Contract Contract has been created."
  Append-ColoredLine $richText Yellow "Please Select the new Contract from the dropdown menu and Edit the Sections as required to add Debtors, Items and Categories (for Sunglasses). Once Contract files have been edited, click the Update Contract to save and load the contract in Micronet"

  Get-ChildItem -Path $PSScriptRoot -Directory | Select Name | ForEach-Object {[void] $lstContract.Items.Add($_.Name)}
  $lstContract.EndUpdate()
})

$CloseButton                        = New-Object system.Windows.Forms.Button
$CloseButton.text                    = "Close"
$CloseButton.width                   = 200
$CloseButton.height                  = 30
$CloseButton.location                = New-Object System.Drawing.Point(220,170)
$CloseButton.Font                    = New-Object System.Drawing.Font('Verdana',10)
$Form.Controls.Add($CloseButton)

$CloseButton.Add_Click({
  $Form.close()


})

$UpdateButton                         = New-Object system.Windows.Forms.Button
$UpdateButton.text                    = "Update Contract"
$UpdateButton.width                   = 200
$UpdateButton.height                  = 30
$UpdateButton.location                = New-Object System.Drawing.Point(220,90)
$UpdateButton.Font                    = New-Object System.Drawing.Font('Verdana',10)
$Form.Controls.Add($UpdateButton)

$UpdateButton.Add_Click({
  $richText.Clear()
  $Contract = $lstContract.SelectedItem
  If (!$Contract) {
    Append-ColoredLine $richText Red "Cotract or details not selected."
  } else
  {
    $categories_file = "$PSScriptRoot\$Contract\categories.csv"
    $items_file = "$PSScriptRoot\$Contract\items.csv"
    $debtors_file = "$PSScriptRoot\$Contract\debtors.csv"

    Append-ColoredLine $richText Green "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss") Setting up New Contract for $Contract"
    Append-ColoredLine $richText Yellow "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss") Getting data for the contract"
    Sleep 2
    $products = Import-CSV $categories_file

    foreach ($p in $products) {

      #If products are in A6 Sinlges price
      $dt | Where-Object {$_.ITM_CAT -eq $p.ITM_CAT } | Select-Object ITM_NO,ITM_CAT,@{Name='ITM_SELL';Expression={$p.ITM_SELL}} | Export-CSV "tmp_$Contract.csv" -NoTypeInformation -Append
    }

    $items = Import-CSV $items_file
    foreach ($i in $items){
      #If products are in A6 Sinlges price
      $dt | Where-Object {$_.ITM_NO -like $i.ITM_NO }  | Select-Object ITM_NO,ITM_CAT,@{Name='ITM_SELL';Expression={$i.ITM_SELL}} | Export-CSV "tmp_$Contract.csv" -NoTypeInformation -Append
    }

    $data = Import-Csv "tmp_$Contract.csv"

    $data | Group-Object -Property ITM_NO | ForEach-Object {
        $currentgroup = $_
        [pscustomobject]@{
            ITM_NO = $currentgroup.Group[0].ITM_NO
            ITM_SELL  = $currentgroup.Group[0].ITM_SELL
            SEQNUM =   $seqnum = $seqnum + 10

        }
        } | Export-csv "tmp_$Contract-seq.csv" -NoTypeInformation
    # Delete old Rolling Contracts
    Sleep 2
    Append-ColoredLine $richText Yellow "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss") Deleting old $Contract Contract"

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

    Sleep 5
    Append-ColoredLine $richText Yellow "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss") Creating new $CONTRACT contract"

    $query = @(
          "INSERT INTO Contract_Header_File (CONTH_NO, CONTH_DES, CONTH_TYPE) VALUES ('$CONTRACT','Contract for $CONTRACT','0')"
            )

    foreach ($q in $query){
      $cmd = new-object System.Data.Odbc.OdbcCommand($q,$conn)
      $cmd.ExecuteNonQuery() | Out-Null
    }

    $debtors = Import-CSV "$debtors_file"
    $seqnum = 0
    foreach ($debtor in $debtors){
      $SEQNUM =  $seqnum + 10
      $DBT_NO = $debtor.DBT_NO
      $q = "INSERT INTO Contract_Application_File (CONTA_NO, CONTA_DBTNO, CONTA_TYPE, CONTA_SEQ) VALUES ('$CONTRACT','$DBT_NO','0','$SEQNUM')"
      $cmd = new-object System.Data.Odbc.OdbcCommand($q,$conn)
      $cmd.ExecuteNonQuery() | Out-Null
    }

    Sleep 5

    $data = Import-CSV "tmp_$contract-seq.csv"
    $data | Group-Object -Property ITM_NO | ForEach-Object {
        $currentgroup = $_
        $ITM_NO = $currentgroup.Group[0].ITM_NO
        $ITM_SELL  = $currentgroup.Group[0].ITM_SELL
        $SEQNUM =   $currentgroup.Group[0].SEQNUM
        Append-ColoredLine $richText Yellow "$(Get-Date -Format "dd/MM/yyyy HH:mm:ss") Adding item $ITM_NO to $CONTRACT Contract"

        $query = @(
                "INSERT INTO Contract_Line_File (CONTL_NO, CONTL_ITMNO, CONTL_TYPE, CONTL_DEFPRICE, CONTL_RETDEF, CONTL_TRADE0, CONTL_TRADE1, CONTL_TRADE2, CONTL_TRADE3, CONTL_TRADE4, CONTL_TRADE5, CONTL_TRADE6, CONTL_TRADE7,
                CONTL_RETAIL0, CONTL_RETAIL1, CONTL_RETAIL2, CONTL_RETAIL3, CONTL_RETAIL4, CONTL_RETAIL5, CONTL_RETAIL6, CONTL_RETAIL7,
                CONTL_SEQ) VALUES ('$CONTRACT','$ITM_NO','0','0','1','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL'
                ,'$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$ITM_SELL','$SEQNUM')"
              )

        foreach ($q in $query){
          $cmd = new-object System.Data.Odbc.OdbcCommand($q,$conn)
          $cmd.ExecuteNonQuery() | Out-Null
          }

        }
    Remove-Item ".\tmp*.csv"

    Append-ColoredLine $richText Blue "Finished creating contract for $Contract"
  }
})

#endregion GUI }
[void] $Form.ShowDialog()
