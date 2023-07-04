<#
.SYNOPSIS
  Welcome Emai
.DESCRIPTION
  Welcome Email for new Users
.NOTES
  Version:        1.0
  Author:         Matthew McKinnon
  Creation Date:  01/03/2021
#>



param (
    [Parameter(Mandatory)]
    [string] $user

 )
$Details = Get-ADUser $User -Property *
$Email = $details.EmailAddress
$FirstName = $Details.FirstName
$ext = $details.ipphone
$displayName = $details.DisplayName
$ddi = $details.officephone
$office = $details.office


#Email Details
$From = "noreply@triotrading.com.au"
$To = $email
$Subject = "Welcome to Trio Brothers Trading."
$SMTPSERVER = "mx1.trio.local"

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
Hi $FirstName
<p style=font-family:verdana>
Welcome to Trio Brothers Trading.
<p style=font-family:verdana>
If you're reading this email, it means you've successfully logged on and have access to your mailbox.
<p style=font-family:verdana>
We have you listed with extension number $ext, which has a direct dial of $ddi. Our Reception phone number is 07 3440 5000.
<p style=font-family:verdana>
Your email address is set to $email, and it will appear as $displayName in most modern email applications. If you'd like to change this, let us know. You can access your email remotely by visiting the address <a href='https://exchange.triotrading.com.au/owa'>https://exchange.triotrading.com.au/owa</a> and logging in with your email address and password you use to login to your computer.
<p style=font-family:verdana>
If you require any IT Support, you can contact us using the support@triotrading.com.au address, You can also log tickets at <a href='https://helpdesk.tbt.net.au/'>https://helpdesk.tbt.net.au/</a>. You'll need to login with your email address password that you use to login to your computer.
<p style=font-family:verdana>
Alternatively, you can call Ext 5020 from your desk phone, or 07 3440 5020 from anywhere else. If you need to report an urgent issue, please call if possible.
<p style=font-family:verdana>
We also have an Internal Wiki that we are slowly building wih guides and useful information. These are available at <a href='https://wiki.tbt.net.au'>https://wiki.tbt.net.au</a>. If you cannot find the information you want here or want something added please let us know.
<p style=font-family:verdana>
We use Slack for internal communication, by now you should have received an inviation email to join. Please make sure you follow the links to join. If you need assistance please let us know. You can join at <a href='https://triotrading.slack.com'>https://triotrading.slack.com</a>
<p style=font-family:verdana>
Once again, welcome to Trio Brothers Trading.
<p style=font-family:verdana>
Thanks,
<p style=font-family:verdana>
Matthew





</body>
</html>
"


  #Add '-UseSSL -Credential $mycredentials' for Office 365 Authentication, Remove for Local Exchange
  Send-MailMessage -To $To -From $From -Subject $Subject -BodyAsHtml -Body $Body -Priority High -SmtpServer $SMTPSERVER

   #logoff $sessionid /server:$server /v
