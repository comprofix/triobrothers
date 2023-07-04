c:\windows\system32\cscript.exe //NOLOGO \\trio.local\data\Apps\Office365\OffScrub_O16msi.vbs CLIENTALL
\\trio.local\data\Apps\Office365\setup.exe /configure \\trio.local\data\Apps\Office365\configure.xml
Unregister-ScheduledTask -TaskName "Office365" -Confirm:$false
Restart-Computer -Force
