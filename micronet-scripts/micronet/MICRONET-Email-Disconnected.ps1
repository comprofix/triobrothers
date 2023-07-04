#Scipt forusers with a disconnected session on the server will be logged off.DESCRIPTION

$server = "micronet";

# Get all RDP sessions
$sessions = query user /server:$server | select -skip 1;

# Loop through each session/line returned
foreach ($line in $sessions) {

    $line = -split $line;

    # Check for missing SessionName field/column
    if ($line.length -eq 8) {

        # Get current session state (column 4)
        $state = $line[3];
        $user = $line[0];

        # Get Session ID (column 3) and current idle time (column 5)
        $sessionid = $line[2];
        $idletime = $line[4];

    } else {

        # Get current session state (column 3)
        $state = $line[2];
        $user = $line[0];

        # Get Session ID (column 2) and current idle time (column 4)
        $sessionid = $line[1];
        $idletime = $line[3];
    }

    # If the session state is Disconnected
    if ($state -eq "Disc") {
      if ($user -eq "opmetrix" -or $user -eq "administrator") {
        logoff $sessionid /server:$server
        exit 0
      }

      #Email Details
      $From = "helpdesk@triotrading.com.au"
      $To = "$user@triotrading.com.au","helpdesk@triotrading.com.au"
      $Subject = $env:computername+": $user Not Logged Out"
      $SMTPSERVER = "mail.trio.local"

      #Email Body
$Body = "
  <!DOCTYPE html>
  <html>
  <head>
  <title>$HTMLMessageSubject</title>
  <style>
    h1.failed {color:red;}
  </style>
  </head>
  <body>
  <h1 class=failed>MICRONET: $user Not Logged Out</h1>
  <p style=font-family:verdana>
    Hi $user,
  <p style=font-family:verdana>
    It appears that you have not logged out of Micronet Correctly.
  <p style=font-family:verdana>
  Can you please make sure you logout of Micronet before you leave.
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
    }
}
