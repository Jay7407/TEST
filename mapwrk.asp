<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Grand Chapter of Texas - Order of the Eastern Star</title>
<meta name="description" content="The Order of the Eastern Star is the largest fraternal organization in the world to which both menand women may belong.">
<meta name="keywords" content="Order of the Eastern Star, Eastern Star, Star, OES,  mason, masonic organization, masonic, ESTARL, fraternal, fraternal organization, Texas Grand Chapter, Grand Chapter OES, Rob Morris">
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <script src="http://maps.google.com/maps?file=api&v=1&key=ABQIAAAAZtImI6l2SqZPxAvHs4gaeBTb3m7wzWZI5ULTZxJubAfk6YOBmRR4RfqdQvW2rh1tV4aPeaSka-dzZw" type="text/javascript"></script>
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/main_menu.jpg">
 <form name="form1" method="post" action="">
  <tr> 
   <td width="76%" rowspan="2">
<script language="JavaScript" src="menu.js"></script> <script language="JavaScript" src="mainmenu.js"></script>
			
			&nbsp;</td>
     </tr>
 </form>
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
            <td colspan="2" valign="top" align="center"><table border="1" cellspacing="0" cellpadding="0">
              <tr>
               <td><div id="map" style="width: 620px; height: 400px"></div></td>
              </tr>
             </table> 
             </td>
           </tr>
          </tbody>
         </table>
									<hr>
         <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr> 
           <td>
            <br> <p align="center" class="header">1503 East Division<br>
             Arlington, Texas 76012</p>
            <p align="center"> 817-265-6263 <br>
             FAX 817-274-5995</p></td>
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
  <td width="341" bgcolor="#990000" class="WhiteText"><div align="center">Copyright � Grand Chapter of Texas<br>
    Order of the Eastern Star</div></td>
  <td bgcolor="#990000" class="footer"><p align="center" class="WhiteText"> 1503 East Division<br>
    Arlington, Texas 76012</p>
   </td>
 </tr>
</table>
</body>
</html>
<script type="text/javascript">
    //<![CDATA[
    
    var map = new GMap(document.getElementById("map"));
    map.addControl(new GSmallMapControl());
    map.centerAndZoom(new GPoint(-097.093987, 32.739895 ), 4);
var point = new GPoint(parseFloat(-097.094987),parseFloat(32.7413));
var marker = new GMarker(point);
map.addOverlay(marker);
GEvent.addListener(marker, "click", function() {
          marker.openInfoWindowHtml("This is where the Grand Office is located.<br>1111 East Division<br>Arlington, Texas 76011");
    });





    //]]>
    </script>
