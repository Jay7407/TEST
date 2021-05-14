<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Contact Us - Grand Chapter of Texas - Order of the Eastern Star</title>
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/main_menu.jpg">
 <form name="form1" method="post" action="">
  <tr> 
   <td width="76%" rowspan="2">
<script language="JavaScript" src="menu.js"></script> <script language="JavaScript" src="mainmenu131.js"></script>
			
			&nbsp;</td>
   </tr>
 </form>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
 <tr> 
  <td width="341" height="198" align="center" valign="top" background="images/main_starl.jpg"><img src="images/spacer.gif" width="332" height="175"></td>
  <td width="100%" align="center" valign="top" background="images/main_menu1.jpg"><p><br>
    <img src="images/main_name.gif" width="324" height="168"></p>
   </td>
 </tr>
 <tr> 
  <td colspan="2" align="center" valign="top"><div align="center">
    <p><%= FormatDateTime(Date, 1) %><span class="header"><br>
     </span><strong><br>
     </strong><span class="header">Contact Us</span></p>
    <p align="left">
    </p>
    </div>
    <table width="98%" border="0" cellspacing="0" cellpadding="5">
     
    <tr> 
     <td valign="top"><p align="left"> 
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

strSMTPServerAddress = "localhost"

'-------------------- Place Carbon Copy e-mail address in the following sting ------------------------

strCCEmailAddress = Request.Form("ToPerson")

'-------------------- Place Blind Copy e-mail address in the following sting -------------------------

strBCCEmailAddress = "yorkja@suddenlink.net" 'Use this string only if you want to send the blind copies of the e-mail yorkja@suddenlink.net

'-----------------------------------------------------------------------------------------------------


'Read in the users e-mail address
strReturnEmailAddress = Request.Form("youremail")


'Initialse strBody string with the body of the e-mail
strBody = "<font size='2' face='Verdana, Arial, Helvetica, sans-serif'>"
strBody = strBody &  "<h4>E-mail sent from form on the Grand Chapter of Texas Web Site</h4>"
strBody = strBody & "<br><br><b>Name: </b>"& Request.Form("yourname")
strBody = strBody & "<br><br><b>Telephone: </b>" & Request.Form("yourphone")
strBody = strBody & "<br><b>E-mail: </b>" & strReturnEmailAddress
strBody = strBody & "<br><br><b>Enquiry: </b><br>" & Replace(Request.Form("message"), vbCrLf, "<br>")
strBody = strBody &  "</font>"


'Check to see if the user has entered an e-mail address and that it is a valid address otherwise set the e-mail address to your own otherwise the e-mail will be rejected
If Len(strReturnEmailAddress) < 5 OR NOT Instr(1, strReturnEmailAddress, " ") = 0 OR InStr(1, strReturnEmailAddress, "@", 1) < 2 OR InStrRev(strReturnEmailAddress, ".") < InStr(1, strReturnEmailAddress, "@", 1) Then
	
	'Set the return e-mail address to your own
	strReturnEmailAddress = strMyEmailAddress
End If	


'Send the e-mail

'Create the e-mail server object
Set objJMail = Server.CreateObject("JMail.SMTPMail")

'Out going SMTP mail server address
objJMail.ServerAddress = strSMTPServerAddress

'Senders email address
objJMail.Sender = strReturnEmailAddress

'Senders name
objJMail.SenderName = Request.Form("yourname")

'Who the email is sent to
objJMail.AddRecipient strMyEmailAddress

'Who the carbon copies are sent to
objJMail.AddRecipientCC strCCEmailAddress

'Who the blind copies are sent to
objJMail.AddRecipientBCC strBCCEmailAddress

'Set the subject of the e-mail
objJMail.Subject = "Grand Chapter Enquiry"


'Set the main body of the e-mail (HTML format)
objJMail.HTMLBody = strBody

'Set the main body of the e-mail (Plain Text format)
'objJMail.Body = strBody

'Importance of the e-mail ( 1 - highest priority, 3 - normal, 5 - lowest)
objJMail.Priority = 3 

'Send the e-mail
'objJMail.ServerAddress = "username:postmaster@grandchapteroftexasoes.org"
objJMail.Execute	
'Close the server object
Set objJMail = Nothing
%>
       The following message was sent:</p>
      <blockquote> 
       <blockquote>
        <p align="left"><%response.write strBody%></p>
       </blockquote>
      </blockquote></td>
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
  <<td width="340" bgcolor="#990000" class="WhiteText"><center>Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</center></td>
    <td  bgcolor="#990000" class="WhiteText">&nbsp;<center>
     8101 Valcasi Drive<br>
    Arlington, Texas 76001</center>&nbsp;</td>
 </tr>
</table>
</body>
</html>
