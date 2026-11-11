<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,chktermcode
	  ' Create a recordset and open it for adding records
chkTerm()
IF chktermcode<>"1" then
	  sSQL = "DELETE FROM pTermcode WHERE sTermcode='" & Request.QueryString("sTermcode") &"'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "pTermcodeall.asp"
else
		response.write "Can't Delete"
End IF


Sub chkTerm()
	Dim Rs2,sSQL2
    Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select  stermcode from pvender where stermcode='" & Request.QueryString("stermcode") & "'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS2.EOF then
			chktermcode=1
		else
			chktermcode=0
		End if
	Rs2.Close
	Set Rs2=nothing
End Sub
%>
