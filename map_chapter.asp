<%
dim dsn1
dim Conn1
dsn1="DBQ=" & Server.Mappath("../Chapters.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open dsn1
 set rs = Conn1.Execute("SELECT * FROM chapters Where Chapter_No = "&request.queryString("ChapterNumber")&";")
If rs.eof and rs.bof Then
  Response.redirect ("Chapter_search.asp")
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Grand Chapter of Texas - Order of the Eastern Star</title>
<meta name="description" content="The Order of the Eastern Star is the largest fraternal organization in the world to which both menand women may belong.">
<meta name="keywords" content="Order of the Eastern Star, Eastern Star, Star, OES,  mason, masonic organization, masonic, ESTARL, fraternal, fraternal organization, Texas Grand Chapter, Grand Chapter OES, Rob Morris">
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

</head>

<body leftmargin="0" topmargin="0" onLoad="loadMap()"> 
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/main_menu.jpg">
  <tr> 
   <td width="76%" rowspan="2">
<script language="JavaScript" src="menu.js"></script> <script language="JavaScript" src="mainmenu131.js"></script>
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
  <td colspan="2" align="center" valign="top"><%= FormatDateTime(Date, 1) %><br>
  </td>
 </tr>
 <tr> 
  <td colspan="2" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
     <td width="5%" height="18" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
      </table>
      <p>&nbsp;</p></td>
     <td width="95%" valign="top"><table border="0" width="98%" cellpadding="5">
       <tr> 
        <td valign="top" style="BORDER-RIGHT: medium none; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; BORDER-BOTTOM: medium none">
         <table width="98%" border="0" cellpadding="2" cellspacing="2">
          <tbody>
           <tr> 
            <td width="65%" align="center" valign="top"><table border="1" cellspacing="0" cellpadding="0">
              <tr> 
              <iframe
  width="400"
  height="420"
  frameborder="0" style="border:0"
  src="https://www.google.com/maps/embed/v1/place?key=AIzaSyCE4LuKEWNMulssSU6D6irKs2FeBM2A848
    &q=(<%=rs("Lat")%>, <%=rs("Lon")%>)">
</iframe>
              </tr>
             </table></td>
            <td width="35%" align="center" valign="top"><p align="left">&nbsp;</p>
             <blockquote>&nbsp;</blockquote>
             <p align="center" class="header">
             <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
               <td width="33%" valign="top">Chapter:</td>
               <td colspan="2"><p><b><%=rs("Chapter_Name")%></b> </p>
                <p><%=Replace(rs("Chapter_Address"), ", ", "<br>")%></p></td>
              </tr>
              <tr> 
               <td>&nbsp;</td>
               <td colspan="2"><%=rs("Time of Meeting")%><br><br></td>
              </tr>
              <tr> 
               <td>Worthy Matron:</td>
               <td width="67%"><%=rs("Worthy_Matron")%></td>
              </tr>
              <tr> 
               <td>Worthy Patron:</td>
               <td><%=rs("Worthy_Patron")%></td>
              </tr>
              <tr> 
               <td>Secretary:</td>
               <td><%=rs("Secretary")%></td>
              </tr>
              <tr> 
               <td>Ph.</td>
               <td><%=rs("Sec_Ph_Number")%></td>
              </tr>
              <tr> 
               <td>&nbsp;</td>
               <td><%=Tmp_Address%></td>
              </tr>
             </table>													
													
													</p>
             <p><a href="chapter_search.asp">Return to Chapter search</a></p></td>
           </tr>
<td width="33%" valign="top">Zoom in or Click on View Large Map, Then Zoom in some and Click on Sat. View in lower Left, then Click on the gold man in the lower Right and click on map this will bring up street view and you can rotate map to see the Building if needed can rotate 360 Deg.</td>           
          </tbody>
         </table>
									<hr>
         <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr> 
           <td>
            <br> <p align="center" class="header">8101 Valcasi Drive<br>
              Suite 101<br>
Arlington, Texas 76001</p>
            <p align="center">817-563-1244<br>
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
  <td width="341" bgcolor="#990000" class="WhiteText"><div align="center">Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</div></td>
  <td bgcolor="#990000" class="footer"><p align="center" class="WhiteText">8101 Valcasi Drive<br>
Arlington, Texas 76001</p>
   </td>
 </tr>
</table>
</body>
</html>
<script type="text/javascript">
    //<![CDATA[

   var WINDOW_HTML = '<div style="width: 12em; style: font-size: small">Grand Chapter Office<br>1111 East Division<br>Arlington, Texas 76011.</div>';

    function loadMap() {
      var map = new GMap(document.getElementById("map"));
      var point = new GPoint(<%=rs("Lon")%>, <%=rs("Lat")%>); 
        map.addControl(new GLargeMapControl()); 
        map.addControl(new GMapTypeControl()); 
        map.centerAndZoom(point, 7); 
						<%	Do while NOT rs.EOF 
						  If Len(rs("Lon"))>0 Then
								response.write rs("Lon")
								%>
      var marker = new GMarker(new GPoint(<%=rs("Lon")%>, <%=rs("Lat")%>));
      map.addOverlay(marker);
						<% 
						End If
						rs.movenext
Loop
rs.close
%>


//      GEvent.addListener(marker, 'click', function() {
//        marker.openInfoWindowHtml(WINDOW_HTML);
//      });
//      marker.openInfoWindowHtml(WINDOW_HTML);
    }
    loadMap();

    //]]>
    </script>
