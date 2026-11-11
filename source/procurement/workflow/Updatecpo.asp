<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->

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

		sSQL="Select sRefno from ppo where cTax='2' and cInvoice='2'"
		  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		  If Not RS2.EOF then
		  	Do While Not RS2.EOF
			sSQL="Select cpo from ppr where sprno ='" & RS2("sRefno")&"'"
			  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
				if Not RS.EOF then
					RS("cpo")="4"
					Rs.Update
				End if
			RS.Close
			RS2.MoveNext
			Loop
		End IF
			Response.write "Update Complete"
%>