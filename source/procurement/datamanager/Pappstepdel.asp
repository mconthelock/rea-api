<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,nstep
	nstep=Request("nstep")
	  ' Create a recordset and open it for adding records
	  sSQL = "DELETE FROM pappstep WHERE sflow='" & Request.QueryString("sflow") &"'" &_
		" and nstep='" & nstep & "'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
if Request("status")="1" then
	  Response.Redirect "pappsteptype.asp?Flow=" & Request.QueryString("sflow")
else
	  Response.Redirect "pappstepall.asp"
End if
%>
