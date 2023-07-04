$ErrorActionPreference = 'Stop'

$filepath      = "\\files\apps\Office2016\Setup.exe"

$packageArgs = @{
  packageName    = 'office'
  fileType       = $fileType
  file           = $filepath
  validExitCodes = @(0, 1223)
  }
Install-ChocolateyInstallPackage @packageArgs
