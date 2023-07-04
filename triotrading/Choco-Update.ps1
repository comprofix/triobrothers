$outdated = choco outdated -r | findstr "false" | %{ $_.Split('|')[0]; }

$outdated
exit

ForEach ($package in $outdated) {
  choco upgrade $package -y --no-progress | Out-File -Append C:\Logs\ChocoUpdate.log
}
$date = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
Rename-Item C:\Logs\ChocoUpdate.log -NewName "ChocoUpdate-$date.log"
Restart-Computer
