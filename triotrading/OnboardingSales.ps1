$R = Read-Host "Enter Area"

$AD = $R + "@*"
$MB = @(get-mailbox | Where-Object {$_.EmailAddresses -like "smtp:$AD"})

$AL = $MB.Alias
$EM = $MB.EmailAddresses


ForEach ($A in $EM)
{
  if ($A -like "smtp:$AD" )
  {
    Write-host "Removing Alias" $A "from" $AL -ForegroundColor Yellow
    Set-Mailbox -Identity $AL -EmailAddresses @{remove="$A"} -EmailAddressPolicyEnabled $False

  }
}
