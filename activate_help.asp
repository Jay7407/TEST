<%
	If request.queryString("ID") = "" Then
   response.redirect "help_form.asp"
 End If
Dim dsn1
Dim Conn1
dsn1="DBQ=" & Server.Mappath("../gethelp.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open dsn1
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open "information", Conn1, 2, 2
RS.Find "ID=" & request.queryString("ID")

RS("Active") = true
RS.update
RS.close
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
        <script language="JavaScript" src="menu.js"></script> <script language="JavaScript" src="mainmenu.js"></script>
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
          </p>
    <p></p>
    </div>
    <table width="98%" border="0" cellspacing="0" cellpadding="5">
    <tr> 
     <td valign="top"><p align="center"> 
              <%		 
Set RS = Server.CreateObject("ADODB.Recordset")
RS.Open "information", Conn1, 2, 2
RS.Find "ID=" & request.queryString("ID")
%>

              <strong>The Following record has been activated.</strong></p>
            <p align="center"><img src="images/bar.jpg" width="420" height="51"></p>
            <table width="56%" border="0" align="center" cellpadding="5" cellspacing="0" class="Border">
              <tr bgcolor="#DFDFDF">
                <td width="23%">Name</td>
                <td width="77%"><%=RS("Name")%></td>
              </tr>
              <tr>
                <td>District</td>
                <td><%=RS("District")%></td>
              </tr>
              <tr bgcolor="#CCCCCC">
                <td bgcolor="#DFDFDF">Chapter #</td>
                <td bgcolor="#DFDFDF"><%=RS("Chapter")%></td>
              </tr>
              <tr>
                <td>E-Mail</td>
                <td><a href="mailto:<%=RS("Email")%>"><%=RS("Email")%></a></td>
              </tr>
              <tr bgcolor="#CCCCCC">
                <td bgcolor="#DFDFDF">Phone Number</td>
                <td bgcolor="#DFDFDF"><%=RS("Phone")%></td>
              </tr>
              <tr>
                <td>Type of Help</td>
                <td><%=RS("Talent")%></td>
              </tr>
            </table>
            </td>
     </tr>
    </table>
<%
RS.close
set RS = nothing
%>

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
  <td width="341" bgcolor="#990000" class="WhiteText"><div align="center">Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</div></td>
  <td bgcolor="#990000">&nbsp;</td>
 </tr>
</table>
</body>
</html>
