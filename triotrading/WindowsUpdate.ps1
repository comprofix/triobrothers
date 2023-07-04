<#
.SYNOPSIS
	Run Windows Updates

.PARAMETER URL
	Run Windows Updates on Servers, Workstations or Single Computer

.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021

.PARAMETER SERVERS
	Flag to Update Servers OU

.PARAMETER WORKSTATIONS
	Flag to Update Workstations OU

.PARAMETER COMPUTERNAME
	Flag to Update Single Computer

.PARAMETER RESTART
	Restart during updates

.EXAMPLE
  .\Windows-Update.ps1 -Servers -Restart
  Update Servers OU and Restart during operation

.EXAMPLE
  .\Windows-Update.ps1 -Servers
  Update Servers OU

.EXAMPLE
  .\Windows-Update.ps1 -Workstations -Restart
  Update Workstations OU and Restart during operation

.EXAMPLE
  .\Windows-Update.ps1 -Workstations
  Update Workstaitons OU

.EXAMPLE
  .\Windows-Update.ps1 -ComputerName upstairs-pc -Restart
  Update Computer and Restart during operation

.EXAMPLE
  .\Windows-Update.ps1 -ComputerName upstairs-pc
  Update Computer Named

#>

param (
	[Parameter(ParameterSetName = 'SERVERS')]
	[switch] $SERVERS,

	[Parameter(ParameterSetName = 'WORKSTATIONS')]
	[switch] $WORKSTATIONS,

	[Parameter(ParameterSetName = 'COMPUTERNAME')]
	[string] $COMPUTERNAME,

	[Parameter(ParameterSetName = 'SERVERS')]
	[Parameter(ParameterSetName = 'WORKSTATIONS')]
	[Parameter(ParameterSetName = 'COMPUTERNAME')]
	[switch] $RESTART


)

if ($SERVERS -eq $true) {
	$computers = Get-ADComputer -Filter * -SearchBase "OU=Servers,OU=Computers,OU=Trio Trading,dc=trio,dc=local" | Where-Object {$_.Name -notmatch 'micronet' -and $_.Name -notmatch 'mx1' -and $_.Name -notmatch 'veeambackup' -and $_.Name -notmatch 'rapp'}
}

if ($WORKSTATIONS -eq $true) {
	$computers = Get-ADComputer -Filter * -SearchBase "OU=Workstations,OU=Computers,OU=Trio Trading,dc=trio,dc=local" | Where-Object {$_.distinguishedname -notmatch 'OU=Laptops*' -and $_.Name -notmatch 'SONIA-PC' -and $_.Name -notmatch 'WAREHOUSE-01' -and $_.Name -notmatch 'ADMIN'}
}

if ($COMPUTERNAME) {
	$computers = Get-ADComputer -Filter * | Where-Object { $_.Name -eq "$COMPUTERNAME" }
}


ForEach ($comp in $computers){
	$ComputerName = $Comp.Name
	Write-Host "Testing Connection to $ComputerName"
	$PingTest = test-connection $ComputerName -Count 1 -ErrorAction SilentlyContinue
	IF ($PingTest) {
		$SHARE = Test-Path -Path "\\$ComputerName\c$\"
		IF (!$SHARE) {
			Write-Host "Creating C$ Share"
			invoke-command -ComputerName $ComputerName -ScriptBlock {net share C$=C:\ }
		}

		IF ($RESTART -eq $true) {
			Restart-Computer -Wait -ComputerName $ComputerName -Force
		}
		$nugetinstall = invoke-command -ComputerName $ComputerName -ScriptBlock {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 ; Install-PackageProvider -Name NuGet -Force}
  	invoke-command -ComputerName $ComputerName -ScriptBlock {[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 ; install-module pswindowsupdate -force -AllowClobber }
  	invoke-command -ComputerName $ComputerName -ScriptBlock {Import-Module PSWindowsUpdate -force}
  	Do {
    	#Reset Timeouts
      $connectiontimeout = 0
      $updatetimeout = 0
  		do{
      	$session = New-PSSession -ComputerName $ComputerName
        "reconnecting remotely to $ComputerName"
        sleep -seconds 10
        $connectiontimeout++
      } until ($session.state -match "Opened" -or $connectiontimeout -ge 10)

      "Checking for new updates available on $ComputerName"
      $updates = invoke-command -session $session -scriptblock {Get-wulist -verbose}
      $updatenumber = ($updates.kb).count
      if ($updates -ne $null){
      	invoke-command -ComputerName $ComputerName -ScriptBlock { Invoke-WUjob -ComputerName localhost -Script "ipmo PSWindowsUpdate; Install-WindowsUpdate -AcceptAll | Out-File C:\PSWindowsUpdate.log" -Confirm:$false -RunNow}
      	sleep -Seconds 30
        do {$updatestatus = Get-Content \\$ComputerName\c$\PSWindowsUpdate.log
        	"Currently processing the following update:"
          Get-Content \\$ComputerName\c$\PSWindowsUpdate.log | select-object -last 1
          sleep -Seconds 10
          $ErrorActionPreference = 'SilentlyContinue'
          $installednumber = ([regex]::Matches($updatestatus, "Installed" )).count
          $Failednumber = ([regex]::Matches($updatestatus, "Failed" )).count
          $ErrorActionPreference = ‘Continue’
          $updatetimeout++
        } until ( ($installednumber + $Failednumber) -eq $updatenumber -or $updatetimeout -ge 720)
        invoke-command -computername $ComputerName -ScriptBlock {Unregister-ScheduledTask -TaskName PSWindowsUpdate -Confirm:$false}
        invoke-command -computername $ComputerName -ScriptBlock {NetSh Advfirewall set allprofiles state off}
        $date = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
				if ($RESTART -eq $true) {
					Restart-Computer -Wait -ComputerName $ComputerName -Force
				}
        Rename-Item \\$ComputerName\c$\PSWindowsUpdate.log -NewName "WindowsUpdate-$date.log"
      }


    } until($updates -eq $null)
    "Windows is now up to date on $ComputerName"
		invoke-command -ComputerName $ComputerName -ScriptBlock {
		net stop wuauserv
		net start wuauserv
	}
	Invoke-Command -ComputerName $ComputerName -ScriptBlock {$updateSession = new-object -com "Microsoft.Update.Session"; $updates=$updateSession.CreateupdateSearcher().Search($criteria).Updates; wuauclt /reportnow }
	}
}
