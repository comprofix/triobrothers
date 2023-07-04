# Description: Boxstarter Script
# Author: https://comprofix.com
#
# Install boxstarter:
# 	. { iwr -useb http://boxstarter.org/bootstrapper.ps1 } | iex; get-boxstarter -Force
# NOTE the "." above is required.
#
# Run this boxstarter by calling the following from **elevated** powershell:
#   example: Install-BoxstarterPackage -PackageName <gist>
# Learn more: http://boxstarter.org/Learn/WebLauncher

# Boxstarter options
$Boxstarter.RebootOk=$true # Allow reboots?
$Boxstarter.NoPassword=$false # Is this a machine with no login password?
$Boxstarter.AutoLogin=$true # Save my password securely and auto-login after a reboot

# Workaround for nested chocolatey folders resulting in path too long error
# Trust PSGallery
Get-PackageProvider -Name NuGet -ForceBootstrap
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# Temporary

Disable-UAC
choco feature enable -n=allowGlobalConfirmation
#choco install Win10-PreConfig --cacheLocation $ChocoCachePath -y
choco install 7zip --cacheLocation $ChocoCachePath -y
choco install googlechrome --cacheLocation $ChocoCachePath -y
choco install notepad2-mod --cacheLocation $ChocoCachePath -y
choco install notepadplusplus --cacheLocation $ChocoCachePath -y
choco install dotnetcore --cacheLocation $ChocoCachePath -y
choco install sysinternals --cacheLocation $ChocoCachePath -y
choco install anydesk.install --cacheLocation $ChocoCachePath -y
choco install ccleaner --cacheLocation $ChocoCachePath -y
choco install vcredist-all --cacheLocation $ChocoCachePath -y
choco install pdftkbuilder --cacheLocation $ChocoCachePath -y
choco install foxitreader --cacheLocation $ChocoCachePath -y
choco install greenshot --cacheLocation $ChocoCachePath -y
choco install slack --cacheLocation $ChocoCachePath -y
choco install TelnetClient -source windowsFeatures -y
choco install NetFX3 -source windowsFeatures -y
choco install office --cacheLocation $ChocoCachePath -y
choco install micronetodbc --cacheLocation $ChocoCachePath -y

# clean up the cache directory
Remove-Item $ChocoCachePath -Recurse

#--- Restore Temporary Settings ---
choco feature disable -n=allowGlobalConfirmation
Enable-MicrosoftUpdate
Install-WindowsUpdate -acceptEula
Enable-UAC
