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


		sSQL="Select * from ppr where sprno='" & Request("sprno") & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				RS("cpo")="2"
				Rs.Update
			End if
			RS.Close

		sSQL="Select * from ppo where spono='" & spono & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS("cstatus")="2"
			Rs.Update
			RS.Close

			sSQL="Select sProdcode from preceiveitem where srecno='" & Request.QueryString("srecno") &"'" 
			  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
				IF Not Rs2.EOF then
					Do While Not RS2.EOF 
						sSQL="Select * from ppoitem where spono='" & spono & "'" &_
						" and sprodcode ='" & RS2("sProdcode") & "'"
						Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
						if Not Rs.EOF then
							Do While Not Rs.EOF
								Rs("cstatus")="1"
								Rs.Update
								RS.MoveNext
							Loop
						End if
						RS.Close
					RS2.MoveNext
					Loop
				End if

		sSQL="Select * from ppr where sprno='" & Request("sPrno") & "'"
 		 Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS.EOF then
			Do while Not RS.EOF	
				RS("cPo")="2"
				RS.Update
			RS.MoveNext
			Loop
		End if
		Rs.Close

	  sSQL = "DELETE FROM pTax WHERE srecno='" & Request.QueryString("srecno") &"'" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL = "DELETE FROM preceive WHERE srecno='" & Request.QueryString("srecno") &"'" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL = "DELETE FROM preceiveitem WHERE srecno='" & Request.QueryString("srecno") &"'" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  Set Rs = Nothing

	  Response.Redirect "finishRec.asp"
%>
