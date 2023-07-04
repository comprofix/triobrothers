$computers = Get-ADComputer -Filter * -SearchBase "OU=Workstations,OU=Computers,OU=Trio Trading,dc=trio,dc=local" | Where-Object {$_.distinguishedname -notmatch 'OU=Laptops*' -and $_.Name -notmatch 'SONIA-PC'}


ForEach ($computer in $computers) {

  $TEST = Test-WSMan $computer.Name -ErrorAction SilentlyContinue

  If ($TEST) {
    $OS = Invoke-Command -ComputerName $computer.Name -ScriptBlock {
      $outdated = choco outdated -r | findstr "false" | %{ $_.Split('|')[0]; }
      ForEach ($package in $outdated) {
        choco upgrade $package -y --no-progress
      }

    }
  }
  Restart-Computer $computer.name -Wait -Force
}
