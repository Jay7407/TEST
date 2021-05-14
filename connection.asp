<%
dim dsn1
dim Conn1
dsn1="DBQ=" & Server.Mappath("/database/oesadmin.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
Set Conn1 = Server.CreateObject("ADODB.Connection")
Conn1.Open dsn1
%>
