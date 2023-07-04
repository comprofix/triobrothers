<#
.SYNOPSIS
  Email Alter Notification
.DESCRIPTION
  Email Alter for Scheduled IT Outages
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021
.PARAMETER SCHEDULED
  Set Alert Type to a Planned Outage
.PARAMETER UNPLANNED
  Set Alert Type to a Unscheduled Outage
.PARAMETER WHAT
  Description of Alert
.PARAMETER WHEN
  Date when event will happen
.PARAMETER START
  Time even will begin
.PARAMETER WHEN
  Time even will end
.PARAMETER WHY
  Short description of why this event is happening
.PARAMETER IMPACT
  What Impact this will have
.PARAMETER COMMENTS
  Any additional comments you want to add
.EXAMPLE
  .\Email_Alert.ps1 -scheduled -WHAT "Trio Servers" -WHEN "08/02/2021" -START "07:00pm AEST" -END "11:00pm AEST" -WHY "Servers require Updates and Reboot" -IMPACT "Mail, Shared Drives, Micronet and other systems will be intermittently unavailable during this time while updates are installed and the servers are reboot."
.EXAMPLE
  .\Email_Alert.ps1 -unplanned -WHAT "Trio Servers" -WHEN "08/02/2021" -START "07:00pm AEST" -END "11:00pm AEST" -WHY "Servers require Updates and Reboot" -IMPACT "Mail, Shared Drives, Micronet and other systems will be intermittently unavailable during this time while updates are installed and the servers are reboot."


#>


param (
    [switch] $UNPLANNED,
    [switch] $SCHEDULED,
    [string] $WHAT,
		[string] $WHEN,
		[string] $START,
		[string] $END,
		[string] $WHY,
		[string] $IMPACT,
		[string] $COMMENTS
 )

# Function to write the HTML Header to the file
Function writeHtmlHeader
{
param($fileName)
$date = ( get-date ).ToString('dd/MM/yyyy')
Add-Content $fileName "<html>"
Add-Content $fileName "<head>"
Add-Content $fileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>"
Add-Content $fileName '<title>DiskSpace Report</title>'
add-content $fileName '<STYLE TYPE="text/css">'
add-content $fileName  "<!--"
add-content $fileName  "td {"
add-content $fileName  "font-family: Verdana;"
add-content $fileName  "font-size: 14px;"
add-content $fileName  "border-top: 1px solid #999999;"
add-content $fileName  "border-right: 1px solid #999999;"
add-content $fileName  "border-bottom: 1px solid #999999;"
add-content $fileName  "border-left: 1px solid #999999;"
add-content $fileName  "padding-top: 0px;"
add-content $fileName  "padding-right: 0px;"
add-content $fileName  "padding-bottom: 0px;"
add-content $fileName  "padding-left: 0px;"
add-content $fileName  "}"
add-content $fileName  "body {"
add-content $fileName  "margin-left: 5px;"
add-content $fileName  "margin-top: 5px;"
add-content $fileName  "margin-right: 0px;"
add-content $fileName  "margin-bottom: 10px;"
add-content $fileName  ""
add-content $fileName  "-->"
add-content $fileName  "</style>"
Add-Content $fileName "</head>"
Add-Content $fileName "<body>"
add-content $fileName  "<table width='50%'>"
add-content $fileName  "<tr>"
add-content $fileName  "<td colspan='1' height='25' width=5% align='left'><img src='.\triologo.png'></td>"
if ($scheduled -eq $true) {
  add-content $fileName  "<td bgcolor='#6699FF' colspan='1' height='25' width=5% align='left'><font face='Verdana' color='#000000' size='5'><strong>Scheduled IT Outage</strong></font></td>"
}
if ($unplanned -eq $true) {
  add-content $fileName  "<td bgcolor='#FF9900' colspan='1' height='25' width=5% align='left'><font face='Verdana' color='#000000' size='5'><strong>Unplanned IT Outage</strong></font></td>"
}

add-content $fileName  "</td>"
add-content $fileName  "</tr>"

}

# Function to write the HTML Header to the file
Function writeTableHeader
{
param($fileName)
Add-Content $fileName "<tr>"
Add-Content $fileName "<td height='50' bgcolor=#999999><b>WHAT:</b></td>"
Add-Content $fileName "<td height='50' bgcolor=#FFFFFF><b>$WHAT</b></td>"
Add-Content $fileName "</tr>"
Add-Content $fileName "<tr>"
Add-Content $fileName "<td height='100' bgcolor=#999999><b>WHEN:</b></td>"
Add-Content $fileName "<td bgcolor=#FFFFFF>"
Add-Content $fileName "<b>Date:</b> $WHEN<p>"
Add-Content $fileName "<b>Start Time:</b> $START<p>"
Add-Content $fileName "<b>End Time:</b> $END<p>"
Add-Content $fileName "</td>"
Add-Content $fileName "</tr>"
Add-Content $fileName "<tr>"
Add-Content $fileName "<td height='50' bgcolor=#999999><b>WHY:</b></td>"
Add-Content $fileName "<td height='50' bgcolor=#FFFFFF>$WHY"
Add-Content $fileName "</tr>"
Add-Content $fileName "<tr>"
Add-Content $fileName "<td height='125' bgcolor=#999999><b>IMPACT:</b></td>"
Add-Content $fileName "<td height='50' bgcolor=#FFFFFF>"
Add-Content $fileName "<tl>"
Add-Content $fileName "<li>$IMPACT</li>"
Add-Content $fileName "</tl>"
Add-Content $fileName "</tr>"
Add-Content $fileName "</table>"
Add-Content $fileName "<p>"
Add-Content $fileName "<font face='Verdana' color='#000000'>"
Add-Content $fileName "$COMMENTS"

}


Function writeHtmlFooter
{
param($fileName)

Add-Content $fileName "</body>"
Add-Content $fileName "</html>"
}

#Global Variaables
$freeSpaceFileName = "FreeSpace.htm"
New-Item -ItemType file $freeSpaceFileName -Force | Out-Null

writeHtmlHeader $freeSpaceFileName
writeTableHeader $freeSpaceFileName


Add-Content $freeSpaceFileName "</table>"
writeHtmlFooter $freeSpaceFileName

#Email Details
$From = "support@triotrading.com.au"
$To = "helpdesk@triotrading.com.au"
$bcc = "office@triotrading.com.au"
if ($unplanned -eq $true) {
  $Subject = "ALERT: Unplanned IT Outage - $WHAT $WHEN $START - $END "
}

if ($scheduled -eq $true) {
  $Subject = "ALERT: Scheduled IT Outage - $WHAT $WHEN $START - $END "
}

#Replace "-Raw" with "| Out-String" when using Powershell 2.0
#
if ($PSVersionTable.PSVersion.Major -eq 2)
{
	$Body = Get-Content $freeSpaceFileName | Out-String
}
else
{
	$Body = Get-Content $freeSpaceFileName -Raw
}

$SMTPSERVER = "mx1.trio.local"

Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -BodyAsHtml -Priority High -SmtpServer $SMTPSERVER -Attachment .\triologo.png -bcc $bcc
rm $freeSpaceFileName
