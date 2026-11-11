<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,chk,sInvno,Rs2,sSQL2,sRecno
	  ' Create a recordset and open it for adding records
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	sInvno=Request.QueryString("sInvno")
	sSQL="Select * from pinvoice where sInvno='" & sInvno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS.EOF then
			sRecno=RS("sRecno")
		End if
		RS.Close

	sSQL="Select * from preceive where sRecno='" & sRecno & "'"

	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL2="Select * from  pTax where sRecno='" & sRecno & "'"
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
	If Not Rs2.EOF then
		chk="1"
	else
		chk="0"
	End if

		If Not RS.EOF then
			if chk="1" then
				RS("cStatus")="T"
			else
				RS("cStatus")="0"
			End IF
			RS.Update				
		End IF
	RS.Close
	  sSQL = "DELETE FROM pInvoice WHERE sInvno='" & Request.QueryString("sInvno") &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  
	  '########## UPDATE CPO, Add by Dermpan on 20-FEB-02 ##########
	  sSql = "update ppr set cPo = '3' where sPrNo = (select sPrNo from preceive where sRecNo = '" & sRecNo & "')"
	  rs.open sSql, conn
	  '########## END UPDATE CPO ##########

	  Set Rs = Nothing
	  Response.Redirect "finishInv.asp"

%>