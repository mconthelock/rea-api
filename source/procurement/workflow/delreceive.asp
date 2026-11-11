<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,chk,sRecby,Rs2,sSQL2,sRecno,spono,sprno
	  ' Create a recordset and open it for adding records
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	sRecby=Request("sRecby")
	sSQL="Select * from pReceive where sRecby='" & sRecby & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS.EOF then
			sRecno=RS("sRecno")
			sPono=Rs("spono")
		End if
		RS.Close

	sSQL="Select * from pPo where spono='" & spono & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS.EOF then
			sprno=RS("sRefno")
			RS("cstatus")="2"
			RS.Update
		End if
		RS.Close
	sSQL="Select * from ppr where sprno='" & sprno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS.EOF then
			RS("cpo")="2"
			RS.Update
		End if
		RS.Close

	  sSQL = "DELETE FROM ptax WHERE sRecno='" & sRecno &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL = "DELETE FROM pInvoice WHERE sRecno='" & sRecno &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL = "DELETE FROM pReceiveitem WHERE sRecno='" & sRecno &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL = "DELETE FROM pReceive WHERE sRecno='" & sRecno &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  Set Rs = Nothing
	  Response.Redirect "delreceive.htm"

%>