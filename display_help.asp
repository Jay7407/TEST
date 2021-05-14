<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Looking For Help - Grand Chapter of Texas - Order of the Eastern Star</title>
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function openWindow(url, name) {
popupWin = window.open(url, name, 'toolbar,scrollbars,resizable,width=450,height=500')
}
// done hiding -->
</script>
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
    <td width="369" height="198" align="center" valign="top" background="images/main_starl.jpg"><img src="images/spacer.gif" width="8" height="175"><img src="images/spacer.gif" width="332" height="175"></td>
    <td width="416" align="center" valign="top" background="images/main_menu1.jpg"><p><br>
    <img src="images/main_name.gif" width="324" height="168"></p></td>
  </tr>
  <tr> 
    <td colspan="2" align="center" valign="top"><div align="center">
        <p><%= FormatDateTime(Date, 1) %><br>
                     </p>
        <p><a href="gethelp.asp">Preform a new Search</a></p>
        <p> <img src="images/bar.jpg" width="420" height="51">        </p>
        <p></p>
      </div>
      <table width="98%" border="0" cellspacing="0" cellpadding="5">
     
        <tr> 
          <td valign="top"><p align="center">
              <%		 
If Request.Form("search")  <>"" Then
dim dsn1
dim Conn1
Dim Found
Found=False
dsn1="DBQ=" & Server.Mappath("../gethelp.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open dsn1
  set rs =Conn1.Execute("SELECT * FROM information WHERE Talent like '%"& Request.Form("search") & "%' AND District="& Request.Form("select") &" ORDER BY Chapter ASC;") 
If not rs.eof and not rs.bof Then
rs.movefirst
do while not rs.eof
Found=True
%>
            <p align="center">&nbsp;</p>
            <table width="425" border="0" align="center" cellpadding="5" cellspacing="0" class="Border">
              <tr>
                <td width="27%" bgcolor="#DFDFDF">Name</td>
                <td width="47%" bgcolor="#DFDFDF"><%=RS("Name")%></td>
                <td width="26%" bgcolor="#DFDFDF">
               <form action="help_email.asp?ID=<%=RS("ID")%>" method="post" target="_blank">
                    <input type=hidden name="RequesterName" value='<%=Request.Form("RequestName")%>'>
                    <input type=hidden name="RequesterEmail" value="<%=Request.Form("RequestEmail")%>">
                    <input type="submit" value="Request Help" name="submit">
								</form></td>
              </tr>
              <tr>
                <td>District</td>
                <td colspan="2"><%=RS("District")%></td>
              </tr>
              <tr>
                <td bgcolor="#DFDFDF">Chapter #</td>
                <td colspan="2" bgcolor="#DFDFDF"><%=RS("Chapter")%></td>
              </tr>
              <tr>
                <td>Type of Help</td>
                <td colspan="2"><%=RS("Talent")%></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <%
rs.moveNext
loop
End If
End If
%>
</p>
<%
RS.close
set RS = nothing
%>
   </p>
   <br>
   </td>
  </tr>
  <tr> 
    <td colspan="2" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="50%" valign="top"> 
            <center>
              <%If NOT Found Then%>
              I am sorry but there is nothing listed for &quot;<%=Request.Form("search")%>&quot; in District <%=Request.Form("select")%> at this time.<br><br><br><br><br><br><br><br><br><br>
              <%End If%>
                
            </center>
            <p>&nbsp;</p></td>
        </tr>
   </table></td>
  </tr>
   
  <td colspan="2" class="footer" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff', endColorStr='#990000', gradientType='0')"><div align="center"><img src="images/spacer.gif" width="8" height="20" align="left">~ 
    <a href="index.asp"> Home</a> | <a href="about_oes.asp">About OES</a> | <a href="membership.asp">Membership</a> ~ </div></td>
  </tr>
  <tr> 
    <td width="369" bgcolor="#990000" class="WhiteText"><div align="center">Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</div></td>
    <td width="416" bgcolor="#990000" class="WhiteText">
        <center>
        1503 W. Division<br>
        Arlington, Texas 76012
       </center></td>
  </tr>
</table>
</body>
</html>
