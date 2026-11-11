<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,status,RS2
	  ' Create a recordset and open it for adding records
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	sSQL="Select sProdcode from ppritem where  sprodcode='" & UCase(Request.QueryString("sprodcode")) &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if RS.EOF then
	  sSQL = "DELETE FROM pproduct WHERE sprodcode='" & Request.QueryString("sprodcode") &"'" &_
		"and sexptype='" & Request.QueryString("sexptype") & "'"
	  'response.write sSql
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs2 = Nothing
		status="del"
     Response.Redirect "Products.asp?status=del"
	else
		Response.write "Can't Delete !!!"
	End if
%>
