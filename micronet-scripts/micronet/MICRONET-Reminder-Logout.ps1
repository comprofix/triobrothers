#Email Details
$From = "helpdesk@triotrading.com.au"
$To = "office@triotrading.com.au"
$Subject = "Please Logout of Micronet before you leave."
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
<h1 class=heading>Please Logout of Micronet</h1>
<p style=font-family:verdana>
Hi Everyone,
<p style=font-family:verdana>
Just a reminder, can you please make sure you logout of Micronet before you leave and when you have finished using Micronet.
<p style=font-family:verdana>
To log out and save work please do the following.
<p style=font-family:verdana>
<ul style=font-family:verdana>
<li style=font-family:verdana>Close all open windows of Micronet and save the changes if prompted.
<li style=font-family:verdana>Click the Start Menu
<li style=font-family:verdana>Click Log off
</ul>
<p>
<img src = '.\micronet-logout-start-menu.png'>


</body>
</html>
"


  #Add '-UseSSL -Credential $mycredentials' for Office 365 Authentication, Remove for Local Exchange
  Send-MailMessage -To $To -From $From -Subject $Subject -BodyAsHtml -Body $Body -Priority High -SmtpServer $SMTPSERVER -Attachment .\micronet-logout-start-menu.png

   #logoff $sessionid /server:$server /v
