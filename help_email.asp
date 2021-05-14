<%
	If request.queryString("ID") = "" Then
   response.redirect "help_form.asp"
 End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Contact Us - Grand Chapter of Texas - Order of the Eastern Star</title>
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<body leftmargin="0" topmargin="0" onLoad="resizeTo(457,470);">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/main_menu.jpg">
  <tr> 
      
    <td width="100%" height="30" rowspan="2" class="header"><center>
        <font color="#FFFFFF">Grand Chapter of Texas Order of the Eastern Star</font>
			</center></td>
   </tr>
</table>
<table width="100%" height="93%" border="0" align="center" cellpadding="0" cellspacing="0">
 
  <tr> 
    <td align="center" valign="top"><div align="center">
        <p>&nbsp;</p>
        <p><%= FormatDateTime(Date, 1) %><br>
          <img src="images/bar.jpg" width="420" height="51">
<br>
          <br>
          <a HREF="javascript:window.close()">Close window</a>

        </p>
        <p>
 
              
          <%
dim dsn1
dim Conn1
dsn1="DBQ=" & Server.Mappath("../gethelp.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open dsn1
  set rs =Conn1.Execute("SELECT * FROM information WHERE ID =  "& request.queryString("ID") & ";") 
rs.movefirst



'Set the response buffer to true so we execute all asp code before sending the HTML to the clients browser
Response.Buffer = True

'Dimension variables
Dim strBody 			'Holds the body of the e-mail
Dim objJMail 			'Holds the mail server object
Dim strMyEmailAddress 		'Holds your e-mail address
Dim strSMTPServerAddress	'Holds the SMTP Server address
Dim strCCEmailAddress		'Holds any carbon copy e-mail addresses if you want to send carbon copies of the e-mail
Dim strBCCEmailAddress		'Holds any blind copy e-mail addresses if you wish to send blind copies of the e-mail
Dim strReturnEmailAddress	'Holds the return e-mail address of the user


'----------------- Place your e-mail address in the following sting ---------------------------------

strMyEmailAddress = "postmaster@grandchapteroftexasoes.org"

'---------- Place the address of the SMTP server you are using in the following sting ---------------

strSMTPServerAddress = "mail.grandchapteroftexasoes.org"

'-------------------- Place Carbon Copy e-mail address in the following sting ------------------------
strCCEmailAddress = RS("email")

'-------------------- Place Blind Copy e-mail address in the following sting -------------------------

'strBCCEmailAddress = "yorkja@amaonline.com" 'Use this string only if you want to send the blind copies of the e-mail

'-----------------------------------------------------------------------------------------------------


'Read in the users e-mail address
strReturnEmailAddress = Request.Form("RequesterEmail")


'Initialse strBody string with the body of the e-mail
strBody = "<font size='2' face='Verdana, Arial, Helvetica, sans-serif'>The following message was sent on " & FormatDateTime(Now, vbLongDate) & " from the OES Help System.<br><br>"
strBody = strBody &  "<br><br><p>"&Request.Form("RequesterName")&" (E-Mail <a  href=mailto:"&Request.Form("RequesterEmail")&">"&Request.Form("RequesterEmail")&"</a> ) has requested your assistance with the following:</p>"

strBody = strBody &  "<table width='75%' border='0' cellspacing='0' cellpadding='0'>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td width='21%'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Name</font></td>"
strBody = strBody &  "    <td width='79%'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&RS("Name")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>District</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&RS("district")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Chapter #</font></td>"
strBody = strBody &  "   <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&RS("Chapter")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>E-Mail</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'><a  href=mailto:"&RS("email")&">"&RS("email")&"</a></font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Phone Number</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&RS("phone")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Type of Help</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&RS("Talent")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "</table>"



'Send the e-mail

'Create the e-mail server object
Set objJMail = Server.CreateObject("JMail.SMTPMail")

'Out going SMTP mail server address
objJMail.ServerAddress = strSMTPServerAddress

'Senders email address
objJMail.Sender = strReturnEmailAddress

'Senders name
objJMail.SenderName = Request.Form("RequesterName")

'Who the email is sent to
objJMail.AddRecipient strMyEmailAddress

'Who the carbon copies are sent to
objJMail.AddRecipientCC strCCEmailAddress

'Who the blind copies are sent to
objJMail.AddRecipientBCC strBCCEmailAddress

'Set the subject of the e-mail
objJMail.Subject = "Grand Chapter Willingness to help others"


'Set the main body of the e-mail (HTML format)
objJMail.HTMLBody = strBody

'Set the main body of the e-mail (Plain Text format)
'objJMail.Body = strBody

'Importance of the e-mail ( 1 - highest priority, 3 - normal, 5 - lowest)
objJMail.Priority = 3 

'Send the e-mail
objJMail.Execute
	
'Close the server object
Set objJMail = Nothing
%>
       <strong>Your request has been sent to <%=RS("Name")%>.</strong></p>
        <p><strong>If they are willing to help you they will contact you shortly.</strong></p>
        <br>
        <br>
      </div>
      <p align="left"><br>
   </p></td>
  </tr>
  
   
  <td class="footer" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff', endColorStr='#990000', gradientType='0')"><div align="center"><img src="images/spacer.gif" width="8" height="20" align="left"></div></td>
  </tr>
  <tr> 
    <td width="340" bgcolor="#990000" class="WhiteText"><center>Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</center></td>
    <td  bgcolor="#990000" class="WhiteText"><center>
     8101 Valcasi Drive<br>
    Arlington, Texas 76001</center></td>
  </tr>
</table>
</body>
</html>
