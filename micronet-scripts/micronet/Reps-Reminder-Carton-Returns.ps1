<#
.SYNOPSIS
  Email Reminder to reps about Carton Returns
.DESCRIPTION
  Email Reminder to reps about Carton Returns
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021
#>

#Email Details
$From = "reception@triotrading.com.au"
$To = "reception@triotrading.com.au", "SalesTeam@triotrading.com.au"
#$To = "matthew@triotrading.com.au"
$Subject = "Carton Returns Reminder"
$SMTPSERVER = "mail.trio.local"

#Email Body
$Body = "
<!DOCTYPE html>
<html>
<head>
<title>$HTMLMessageSubject</title>
<style>
h1.heading {color:green;}
</style>
</head>
<body>
<p style=font-family:verdana>
Hi Everyone,
<p style=font-family:verdana>
Just a reminder, can you please remember to email me with the quantity of how many cartons you are returning this week before Monday morning.
<p style=font-family:verdana>
Thanks
</body>
</html>
"
Send-MailMessage -To $To -From $From -Subject $Subject -BodyAsHtml -Body $Body -Priority High -SmtpServer $SMTPSERVER
