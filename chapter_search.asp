<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Grand Chapter of Texas - Order of the Eastern Star</title>
<meta name="description" content="The Order of the Eastern Star is the largest fraternal organization in the world to which both menand women may belong.">
<meta name="keywords" content="Order of the Eastern Star, Eastern Star, Star, OES,  mason, masonic organization, masonic, ESTARL, fraternal, fraternal organization, Texas Grand Chapter, Grand Chapter OES, Rob Morris">
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script type="text/javascript" src="js/fValidate.config.js"></script>
<script type="text/javascript" src="js/fValidate.core.js"></script>
<script type="text/javascript" src="js/fValidate.lang-enUS.js"></script>
<script type="text/javascript" src="js/fValidate.validators.js"></script>
<script type="text/javascript" src="js/fValidate.validators.js"></script>
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/main_menu.jpg">
  <tr> 
    <td width="76%" height="30" rowspan="2">
<script language="JavaScript" src="menu.js"></script> <script language="JavaScript" src="mainmenu131.js"></script>
			
			&nbsp;</td>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
 <tr> 
  <td width="341" height="198" align="center" valign="top" background="images/main_starl.jpg"><img src="images/spacer.gif" width="332" height="175"></td>
  <td width="100%" align="center" valign="top" background="images/main_menu1.jpg"><p><br>
    <img src="images/main_name.gif" width="324" height="168"></p>
   </td>
 </tr>
 <tr> 
  <td colspan="2" align="center" valign="top"><%= FormatDateTime(Date, 1) %><br>
  </td>
 </tr>
 <tr> 
  <td colspan="2" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
     <td width="5%" height="18" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
       <tr> 
        <td width="46%"><div align="center"></div>
         <div align="center"></div></td>
        <td width="54%"><div align="center"></div>
         <div align="center"></div></td>
       </tr>
      </table>
      
     <p>&nbsp;</p></td>
     <td width="95%" valign="top"><table border="0" width="98%" cellpadding="5">
       <tr> 
        <td valign="top" style="BORDER-RIGHT: medium none; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: medium none"> 
								<form name="form2" method="post" action="Chapter_search.asp" onSubmit="return validateForm( this, 0, 1, 0, 0, 15 );">
          <table width="98%" border="0" cellpadding="2" cellspacing="2">
           <tbody>
            <tr> 
             <td width="13%" valign="top"><br> </td>
             <td width="87%" valign="top"><br>
              Please enter the city or part of the address that you wish to find a Chapter for.<br>
              Note: To list <strong>ALL</strong> Chapters please enter the word <strong>Texas</strong>.</td>
            </tr>
            <tr> 
             <td valign="top">&nbsp;</td>
             <td width="87%" valign="top"><input type="text" name="Address" alt="blank"> <input type="submit" name="Submit2" value="Submit">
              <input name="DoSearch" type="hidden" value="DoIt"> </td>
            </tr>
            <tr> 
             <td colspan="2" valign="top">&nbsp;</td>
            </tr>
           </tbody>
          </table>
         </form>
         
         <p>&nbsp; 
          <% 
If Request.Form	("Dosearch") = "DoIt" Then
dim dsn1, Tmp_Address
dim Conn1
dsn1="DBQ=" & Server.Mappath("../Chapters.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open dsn1
  set rs =Conn1.Execute("SELECT * FROM chapters WHERE Chapter_Address like '%"& Request.Form("Address") & "%'  ORDER BY Chapter_No ASC;") 
If not rs.eof and not rs.bof Then
rs.movefirst
do while not rs.eof
Tmp_Address = Replace(rs("Chapter_Address"), ", ", "<br>")

%>

         <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
           <td width="17%">Chapter:</td>
           <td colspan="2"><b><%=rs("Chapter_Name")%></b> (District <%=rs("Dist")%>)</td>
           <td width="23%" valign="top">&nbsp;</td>
           <td width="1%" rowspan="4">
											<%If Len(rs("Lat")) > 1 Then
											%>
											<a href="map_chapter.asp?ChapterNumber=<%=rs("Chapter_No")%>">Map</a>
											<%End If%>
											</td>
          </tr>
          <tr> 
           <td>Worthy Matron:</td>
           <td width="31%"><%=rs("Worthy_Matron")%></td>
           <td width="28%" rowspan="3" valign="top"><%=Tmp_Address%><br><br>Sec. Ph. <%=rs("Sec_Ph_Number")%></td>
           <td width="23%" rowspan="3" valign="top"><%=rs("Time of Meeting")%><br><br>Sec.Email <%=rs("phone")%></td>
          </tr>
          <tr> 
           <td>Worthy Patron:</td>
           <td><%=rs("Worthy_Patron")%></td>
          </tr>
          <tr> 
           <td>Secretary:</td>
           <td><%=rs("Secretary")%></td>
          </tr>
         </table>
         <hr>
         <div align="center">
          <%
rs.moveNext
loop
End If
Else
%>
          <center><img src="images/Txflg.gif" width="533" height="561"> </center>
          <%
End If
%></p>
          </div>
         <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr> 
           <td><br> <p align="center" class="header">8101 Valcasi Drive<br>Suite 101<br>
               Arlington, Texas 76001</p>
             <p align="center"> 817-563-1244<br>
             FAX 817-563-1701</p></td>
          </tr>
         </table>
         <p align="center">&nbsp;</p>
         </td>
       </tr>
      </table></td>
    </tr>
   </table></td>
 </tr>
 <t> 
  <td colspan="2" class="footer" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff', endColorStr='#990000', gradientType='0')"><div align="center"><img src="images/spacer.gif" width="8" height="20" align="left">~ 
    <a href="index.asp"> Home</a> | <a href="about_oes.asp">About OES</a> | <a href="membership.asp">Membership</a> ~ </div></td>
 </tr>
 <tr> 
  <td width="340" bgcolor="#990000" class="WhiteText"><center>Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</center></td>
    <td  bgcolor="#990000" class="WhiteText">&nbsp;<center>
     8101 Valcasi Drive<br>
    Arlington, Texas 76001</center>&nbsp;
   </td>
 </tr>
</table>
</body>
</html>
