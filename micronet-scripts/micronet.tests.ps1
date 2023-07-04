#Describe 'Micronet Services' {
#    It 'MicronetLLS28 should be running' {
#
#		$server = "micronet"
#    $status = (Invoke-Command -ComputerName $server -ScriptBlock { Get-Service -ServiceName MicronetLLS28}).Status
#		$status | Should -Be 'Running'
#    }
#}





Describe 'Test Connection to Micronet' {
  It 'Micronet should be pingable' {
    Test-Connection -ComputerName micronet.trio.local -Quiet -Count 1 | Should -Be $true
  }
}
