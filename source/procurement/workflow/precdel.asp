<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,spono,RS2
	  ' Create a recordset and open it for adding records
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		spono=Request("spono")
	  sSQL = "DELETE FROM planreceive WHERE ID='" & Request.QueryString("sprecno") &"'" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL = "DELETE FROM planreceiveitem WHERE ID='" & Request.QueryString("sprecno") &"'" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  Set Rs = Nothing

	  Response.Redirect "finishplanRec.asp"
%>
