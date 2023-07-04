$computers = Get-ADComputer -Filter * -SearchBase "OU=Workstations,OU=Computers,OU=Trio Trading,dc=trio,dc=local" | Where-Object {$_.distinguishedname -notmatch 'OU=Laptops*' -and $_.Name -notmatch 'ADMIN' }

ForEach ($computer in $computers) {

  #$TEST = Test-WSMan $computer.Name -ErrorAction SilentlyContinue
  $Name = $Computer.Name
  $TEST = Test-Path -Path \\$Name\c$\Scripts

  If ($TEST) {
    Write-Host $computer.Name "Path is available" -ForegroundColor Green

  }
  Else
  {
    Write-Host $Name "is not available. Creating" -ForegroundColor Red
    New-Item -Path \\$Name\c$\ -Name Scripts -Force
  }
}
