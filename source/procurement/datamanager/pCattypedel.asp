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
	sSQL="Select sCattype from pproduct where sCattype='" & Request.QueryString("sCattype") &"'"
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if RS2.EOF then
	  sSQL = "DELETE FROM pCattype WHERE sCattype='" & Request.QueryString("sCattype") &"'"
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "pCattypeall.asp"
	else
		response.write "<br><center><font color=red size=+2>Can not Delete this Group in Expense Type use in Product !!!</font></br></center>"
	end if
%>
