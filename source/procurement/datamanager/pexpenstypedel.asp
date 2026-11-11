<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,RS2
	  ' Create a recordset and open it for adding records
	sSQL="Select sExptype from pproduct where sexptype='" & Request.QueryString("sexptype") &"'"
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if RS2.EOF then
	  sSQL = "DELETE FROM pexpenstype WHERE sexptype='" & Request.QueryString("sexptype") &"'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "pexpenstypeall.asp"
	else
		response.write "<br><center><font color=red size=+2>Can not Delete this Expense Type use in Product !!!</font></br></center>"
	end if
	  Set Rs2 = Nothing
%>
