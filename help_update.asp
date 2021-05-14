<%
	If request("Name") = "" Then
   response.redirect "help_form.asp"
 End If
Dim dsn1
Dim Conn1
Dim UserID
dsn1="DBQ=" & Server.Mappath("../gethelp.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open dsn1
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open "information", Conn1, 2, 2
RS.addnew

RS("Name") = request("Name")
RS("District") = request("District")
RS("Chapter") = request("Chapter")
RS("Email") = request("Email")
RS("Phone") = request("Phone")
RS("Talent") = request("Talent")
RS("Active") = false

RS.update
 'set RS = Conn1.Execute("SELECT * FROM information Where Name = "&Request.Form("Name")&" AND Talent="&Request.Form("Talent")&" ;")
RS.Find "Name='" & request("Name") & "'"
 UserID = RS("ID")
RS.close
set RS = nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Contact Us - Grand Chapter of Texas - Order of the Eastern Star</title>
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/main_menu.jpg">
  <tr> 
      <td width="100%" height="30" rowspan="2">
        <script language="JavaScript" src="menu.js"></script> <script language="JavaScript" src="mainmenu131.js"></script>
			&nbsp;</td>
   </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
 <tr> 
  <td width="341" height="198" align="center" valign="top" background="images/main_starl.jpg"><img src="images/spacer.gif" width="8" height="175"><img src="images/spacer.gif" width="332" height="175"></td>
  <td width="100%" align="center" valign="top" background="images/main_menu1.jpg"><p><br>
    <img src="images/main_name.gif" width="324" height="168"></p>
   </td>
 </tr>
 <tr> 
  <td colspan="2" align="center" valign="top"><div align="center">
        <p><%= FormatDateTime(Date, 1) %><br>
          <img src="images/need_help.jpg" width="259" height="259" class="Border"></p>
    <p></p>
    </div>
    <table width="98%" border="0" cellspacing="0" cellpadding="5">
     
    <tr> 
     <td valign="top"><p align="center"> 
              <%
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
strCCEmailAddress = "hub@ftwweb.com"

'-------------------- Place Blind Copy e-mail address in the following sting -------------------------

'strBCCEmailAddress = "yorkja@amaonline.com" 'Use this string only if you want to send the blind copies of the e-mail

'-----------------------------------------------------------------------------------------------------


'Read in the users e-mail address
strReturnEmailAddress = Request.Form("email")


'Initialse strBody string with the body of the e-mail
strBody = "<font size='2' face='Verdana, Arial, Helvetica, sans-serif'>The following message was sent on " & FormatDateTime(Now, vbLongDate) & " from the OES Help form<br><br><br>"
strBody = strBody &  "<table width='75%' border='0' cellspacing='0' cellpadding='0'>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td width='21%'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Name</font></td>"
strBody = strBody &  "    <td width='79%'><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&Request.Form("Name")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>District</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&Request.Form("district")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Chapter #</font></td>"
strBody = strBody &  "   <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&Request.Form("Chapter")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>E-Mail</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'><a  href=mailto:"&Request.Form("email")&">"&Request.Form("email")&"</a></font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Phone Number</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&Request.Form("phone")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "  <tr>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>Type of Help</font></td>"
strBody = strBody &  "    <td><font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"&Request.Form("Talent")&"</font></td>"
strBody = strBody &  "  </tr>"
strBody = strBody &  "</table>"

strBody = strBody &  "<p><font face='Verdana, Arial, Helvetica, sans-serif'>If you believe this request is valid you may <a href='http://www.grandchapteroftexasoes.org/activate_help.asp?ID="&UserID&"'>activate it now</a>.</font></p>"


'Send the e-mail

'Create the e-mail server object
Set objJMail = Server.CreateObject("JMail.SMTPMail")

'Out going SMTP mail server address
objJMail.ServerAddress = strSMTPServerAddress

'Senders email address
objJMail.Sender = strReturnEmailAddress

'Senders name
objJMail.SenderName = "OES WEB Server"

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
       <strong>Thank you for your willingness to help our organization grow and to help your Sisters and Brothers.</strong></p>
            <p align="center"><strong>A message has been sent to our Webmaster to validate your information and activate it.</strong></p>
      </td>
     </tr>
     <tr> 
      <td>&nbsp; </td>
     </tr>
    </table>
   <p align="left"><br>
   </p>
   </td>
 </tr>
 <tr> 
  <td colspan="2" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
     <td height="18" valign="top"> <br> <p>&nbsp;</p></td>
    </tr>
   </table></td>
 </tr>
   <td colspan="2" class="footer" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff', endColorStr='#990000', gradientType='0')"><div align="center"><img src="images/spacer.gif" width="8" height="20" align="left">~ 
    <a href="index.asp"> Home</a> | <a href="about_oes.asp">About OES</a> | <a href="membership.asp">Membership</a> ~ </div></td>
 </tr>
 <tr> 
  <td width="340" bgcolor="#990000" class="WhiteText"><center>Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</center></td>
    <td  bgcolor="#990000" class="WhiteText">&nbsp;<center>
     8101 Valcasi Drive<br>
    Arlington, Texas 76001</center>&nbsp;</td>
 </tr>
</table>
</body>
</html>
