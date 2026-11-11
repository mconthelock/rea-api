<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<%
	' Author: worakrit yodthonglert
	' Version: 0.01
%>	
<%
  Dim sSQL,RS
  sSQL = "SELECT * FROM AEMPLOYEE"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
  While not RS.EOF
%>
sEmpno : <%=RS("sEmpno")%><br>
sDivc : <%=RS("sDivc")%><br>
sDepc : <%=RS("sDepc")%><br>
sSecc : <%=RS("sSecc")%><br>
sPO : <%=RS("sPo")%><br>
sName : <%=RS("sName")%><br>
sSurname : <%=RS("sSurname")%><br>
sTName : <%=RS("sTName")%><br>
sTSurname : <%=RS("sTSurname")%><br>
sLogin : <%=RS("sLogin")%><br>
sPassword : <%=RS("sPassword")%><br>
<%
  RS.MoveNext 
  wend
  RS.Close
  Set RS = Nothing
%>	
