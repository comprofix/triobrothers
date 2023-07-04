Describe 'Test Connection to File Server' {
  It 'File Server should be pingable' {
    Test-Connection -ComputerName filesrv.trio.local -Quiet -Count 1 | Should -Be $true
  }
}
