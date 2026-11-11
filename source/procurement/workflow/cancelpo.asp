<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,spono,Color,sprno,sprodcode
	spono=Request("spono")
	sprno=Request("sprno")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
		sSQL="Select * from ppo where spono='" & spono & "'"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		RS("cstatus")="9" 
		Rs.Update
		Rs.Close

		sSQL="Delete from preceiveitem where srecno in" &_
		" (Select sRecno from preceive where spono='" & spono & "')"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from preceive where spono='" & spono & "'"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Select * from ppoitem where spono='" & spono & "'"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not Rs.EOF then 
			Do While Not RS.EOF
				sprodcode=RS("sProdcode")
				sSQL2="Select * from ppritem where sprno='" & sprno & "'" &_
				" and sprodcode='" & sprodcode & "'"
			  		Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
					RS2("cStatus")="0"
					Rs2("sUsefor")= "Reject By " & Session("sEmpno")
					Rs2.Update
					Rs2.Close
				sSQL2="Select * from ppritem2 where sprno='" & sprno & "'" &_
				" and sprodcode='" & sprodcode & "'"
			  		Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
				if Not RS2.EOF then
					RS2("cStatus")="0"
					Rs2("sUsefor")= "Reject By " & Session("sEmpno")
					Rs2.Update
				end if
					Rs2.Close
			Rs.MoveNext
			Loop
		End if
		Set Rs=Nothing
		response.redirect "poprint.asp"
%>