<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->

<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs
	  ' Create a recordset and open it for adding records
	  sSQL = "DELETE FROM ppo WHERE spono='" & Request.QueryString("spono") &"'"
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "ppoall.asp"
%>
