<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs
	  ' Create a recordset and open it for adding records
	  sSQL = "DELETE FROM pSection WHERE sSeccode='" & Request.QueryString("sSeccode") &"'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing

	  Response.Redirect "psectionall.asp"
%>
