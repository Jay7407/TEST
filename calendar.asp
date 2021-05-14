<!--#include file="calendar/dsn.asp"-->
<!--#include file="calendar/body.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Calendar of Events - Grand Chapter of Texas - Order of the Eastern Star</title>
<link href="styles.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/main_menu.jpg">
  <form name="form1" method="post" action="">
    <tr> 
      <td width="76%" height="30">
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
  <td colspan="2" align="center" valign="top"><div align="center"> 
    <p><%= FormatDateTime(Date, 1) %><span class="header"><br>
     </span><strong><br>
     </strong><span class="header"> 
     <%

navmonth = request.querystring("month")
navyear = request.querystring("year")

If navmonth = "" Then
	navmonth = Month(Date)
End If

If navyear = "" Then
	navyear = Year(Date)
End If

firstday = Weekday(CDate(navmonth & "/" & 1 & "/" & navyear))


leapTestNumbers = navyear / 4
leapTest = (leapTestNumbers) - Round(leapTestNumbers)
If navmonth = 2 Then
	If leapTest <> 0 Then
		lastDate = 28
	Else
		lastDate = 29
	End If
ElseIf ((navmonth = 4) OR (navmonth = 6) OR (navmonth = 9) OR (navmonth = 11)) Then
	lastDate = 30
Else
	lastDate = 31
End If

lastMonth = navmonth - 1
lastYear = navyear
If lastMonth < 1 Then
	lastMonth = 12
	lastYear = lastYear - 1
End If

nextMonth = navmonth + 1
nextYear = navyear
If nextMonth >12 Then
	nextMonth = 1
	nextYear = nextYear + 1
End If


dateCounter = 1
weekCount = 1

DateEnd = lastDate
DateBegin = firstDate

Weekday_Font = RSBODY("Weekday_Font")
Weekday_Font_Size = RSBODY("Weekday_Font_Size")
Weekday_Font_Color = RSBODY("Weekday_Font_Color")
Weekday_Background_Color = RSBODY("Weekday_Background_Color")
Date_Font = RSBODY("Date_Font")
Date_Font_Size = RSBODY("Date_Font_Size")
Date_Font_Color = RSBODY("Date_Font_Color")
Date_Background_Color = RSBODY("Date_Background_Color")
Event_Font = RSBODY("Event_Font")
Event_Font_Size = RSBODY("Event_Font_Size")
Event_Font_Color = RSBODY("Event_Font_Color")
Cell_Width = RSBODY("Cell_Width")
Cell_Height = RSBODY("Cell_Height")
Cell_Background_Color = RSBODY("Cell_Background_Color")
Border_Size = RSBODY("Border_Size")
Border_Color = RSBODY("Border_Color")
%>
     <%=RSBODY("Header")%> </span></p>
   </div>
   <table border="0" cellpadding="0" cellspacing="0" align="center">
    <tr> 
     <td align="center"> <table border="0" cellpadding="2" cellspacing="0" width="100%">
       <tr> 
        <td align="left" valign="bottom"><font face="verdana" size="2"><b><a href="calendar/index.asp?month=<%=lastMonth%>&year=<%=lastYear%>">
         <%If (RSBODY("Abbreviate_Months") = True) Then%>
         <%=MonthName(lastMonth, true)%>
         <%Else%>
         <%=MonthName(lastMonth)%>
         <%End If%>
         </a></b></font></td>
        <td align="center" valign="bottom"><font face="verdana" size="4"><b>
         <%If (RSBODY("Abbreviate_Months") = True) Then%>
         <%=MonthName(navMonth, true)%>
         <%Else%>
         <%=MonthName(navMonth)%>
         <%End If%>
         &nbsp;<%=navyear%></b></font></td>
        <td align="right" valign="bottom"><font face="verdana" size="2"><b><a href="calendar/index.asp?month=<%=nextMonth%>&year=<%=nextYear%>">
         <%If (RSBODY("Abbreviate_Months") = True) Then%>
         <%=MonthName(nextMonth, true)%>
         <%Else%>
         <%=MonthName(nextMonth)%>
         <%End If%>
         </a></b></font></td>
       </tr>
      </table></td>
    </tr>
    <tr> 
     <td> <table border="0" cellpadding="0" cellspacing="0" align="center">
       <tr> 
        <td bgcolor="<%=Border_Color%>"> <table border="0" cellpadding="2" cellspacing="<%=Border_Size%>">
          <tr> 
           <td width="<%=Cell_Width%>" align="center" bgcolor="<%=Weekday_Background_Color%>"><font face="<%=Weekday_Font%>" size="<%=Weekday_Font_Size%>" color="<%=Weekday_Font_Color%>"><b>
            <%If (RSBODY("Abbreviate_Days") = True) Then%>
            Sun
            <%Else%>
            Sunday
            <%End If%>
            </b></font></td>
           <td width="<%=Cell_Width%>" align="center" bgcolor="<%=Weekday_Background_Color%>"><font face="<%=Weekday_Font%>" size="<%=Weekday_Font_Size%>" color="<%=Weekday_Font_Color%>"><b>
            <%If (RSBODY("Abbreviate_Days") = True) Then%>
            Mon
            <%Else%>
            Monday
            <%End If%>
            </b></font></td>
           <td width="<%=Cell_Width%>" align="center" bgcolor="<%=Weekday_Background_Color%>"><font face="<%=Weekday_Font%>" size="<%=Weekday_Font_Size%>" color="<%=Weekday_Font_Color%>"><b>
            <%If (RSBODY("Abbreviate_Days") = True) Then%>
            Tues
            <%Else%>
            Tuesday
            <%End If%>
            </b></font></td>
           <td width="<%=Cell_Width%>" align="center" bgcolor="<%=Weekday_Background_Color%>"><font face="<%=Weekday_Font%>" size="<%=Weekday_Font_Size%>" color="<%=Weekday_Font_Color%>"><b>
            <%If (RSBODY("Abbreviate_Days") = True) Then%>
            Wed
            <%Else%>
            Wednesday
            <%End If%>
            </b></font></td>
           <td width="<%=Cell_Width%>" align="center" bgcolor="<%=Weekday_Background_Color%>"><font face="<%=Weekday_Font%>" size="<%=Weekday_Font_Size%>" color="<%=Weekday_Font_Color%>"><b>
            <%If (RSBODY("Abbreviate_Days") = True) Then%>
            Thurs
            <%Else%>
            Thursday
            <%End If%>
            </b></font></td>
           <td width="<%=Cell_Width%>" align="center" bgcolor="<%=Weekday_Background_Color%>"><font face="<%=Weekday_Font%>" size="<%=Weekday_Font_Size%>" color="<%=Weekday_Font_Color%>"><b>
            <%If (RSBODY("Abbreviate_Days") = True) Then%>
            Fri
            <%Else%>
            Friday
            <%End If%>
            </b></font></td>
           <td width="<%=Cell_Width%>" align="center" bgcolor="<%=Weekday_Background_Color%>"><font face="<%=Weekday_Font%>" size="<%=Weekday_Font_Size%>" color="<%=Weekday_Font_Color%>"><b>
            <%If (RSBODY("Abbreviate_Days") = True) Then%>
            Sat
            <%Else%>
            Saturday
            <%End If%>
            </b></font></td>
          </tr>
          <tr> 
           <% Do while weekCount <= 7
dateSelect = navmonth & "/" & dateCounter & "/" & navyear %>
           <% If (weekCount < firstDay) OR (dateCounter > lastDate) Then %>
           <td height="<%=Cell_Height%>" bgcolor="<%=Cell_Background_Color%>"><img src="calendar/im/clear.gif" height="1" width="1"></td>
           <% else %>
           <td height="<%=Cell_Height%>" bgcolor="<%=Date_Background_Color%>" valign="top"><a href="calendar/date.asp?date=<%=dateSelect%>"><font face="<%=Date_Font%>" size="<%=Date_Font_Size%>" color="<%=Date_Font_Color%>"><b><%=dateCounter%></b></font><font face="<%=Event_Font%>" size="<%=Event_Font_Size%>" color="<%=Event_Font_Color%>"> 
            <%
Set RSEVENT = Server.CreateObject("ADODB.RecordSet")
RSEVENT.Open "SELECT * FROM Events", Conn, 1, 3

Do while NOT RSEVENT.EOF
rsdate = RSEVENT("Date")
If (Day(rsdate) = dateCounter) AND (Month(rsdate) = CInt(navmonth)) AND (Year(rsdate) = CInt(navyear)) Then%>
            <br>
            <%If RSBODY("Event_Display") = True Then%>
            <%=RSEVENT("Category")%> 
            <%Else%>
            <%=RSEVENT("Event_Name")%> 
            <% End If

End If
RSEVENT.movenext
Loop
RSEVENT.close
%>
            </font></a></td>
           <% 
dateCounter = dateCounter + 1
end if

weekCount = weekCount + 1
Loop
weekCount = 1
%>
          </tr>
          <% Do while dateCounter <= lastDate %>
          <tr> 
           <% Do while weekCount <= 7
dateSelect = navmonth & "/" & dateCounter & "/" & navyear %>
           <% If dateCounter > lastDate Then %>
           <td height="<%=Cell_Height%>" bgcolor="<%=Cell_Background_Color%>"><img src="calendar/im/clear.gif" height="1" width="1"></td>
           <% else %>
           <td height="<%=Cell_Height%>" bgcolor="<%=Date_Background_Color%>" valign="top"><a href="calendar/date.asp?date=<%=dateSelect%>"><font face="<%=Date_Font%>" size="<%=Date_Font_Size%>" color="<%=Date_Font_Color%>"><b><%=dateCounter%></b></font><font face="<%=Event_Font%>" size="<%=Event_Font_Size%>" color="<%=Event_Font_Color%>"> 
            <%
Set RSEVENT = Server.CreateObject("ADODB.RecordSet")
RSEVENT.Open "SELECT * FROM Events", Conn, 1, 3
Do while NOT RSEVENT.EOF
rsdate = RSEVENT("Date")
If (Day(rsdate) = dateCounter) AND (Month(rsdate) = CInt(navmonth)) AND (Year(rsdate) = CInt(navyear)) Then%>
            <br>
            <%If RSBODY("Event_Display") = True Then%>
            <%=RSEVENT("Category")%> 
            <%Else%>
            <%=RSEVENT("Event_Name")%> 
            <% End If

End If
RSEVENT.movenext
Loop
RSEVENT.close
%>
            </font></a></td>
           <% 
dateCounter = dateCounter + 1
end if

weekCount = weekCount + 1
Loop
weekCount = 1
%>
          </tr>
          <% Loop %>
         </table></td>
       </tr>
      </table></td>
    </tr>
    <!-- Begin Ocean12 Technologies Copyright Notice -->
    <!-- THIS CODE MUST NOT BE CHANGED -->
    <tr> 
     <td align="center">&nbsp;</td>
    </tr>
    <!-- End Ocean12 Technologies Copyright Notice -->
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
 <t> 
  <td colspan="2" class="footer" style="filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr='#ffffff', endColorStr='#990000', gradientType='0')"><div align="center"><img src="images/spacer.gif" width="8" height="20" align="left">~ 
    <a href="calendar/index.asp"> Home</a> | <a href="calendar/about_oes.asp">About OES</a> | <a href="calendar/membership.asp">Membership</a> ~ </div></td>
 </tr>
 <tr> 
  <td width="341" bgcolor="#990000" class="WhiteText"><div align="center">Copyright © Grand Chapter of Texas<br>
    Order of the Eastern Star</div></td>
  <td bgcolor="#990000">&nbsp;</td>
 </tr>
</table>
</body>
</html>
<!--#include file="calendar/dsn2.asp"-->