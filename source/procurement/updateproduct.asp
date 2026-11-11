<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->

<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,sProdcode,sProdname,RS2,str
	  ' Create a recordset and open it for adding records
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		sProdcode=Request("sProdcode")
		sProdname=Request("sProdname")

		sSQL="Select * from pproduct where sprodcode ='" & sProdcode & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					RS("sEprodname")=sProdname
					RS.update
					RS.MoveNext
				Loop
			End if
			RS.Close

		sSQL="Select * from ppritem where sprodcode='" & sprodcode & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					RS("sprodname")=sprodname
					RS.Update
					RS.MoveNext
				Loop
			End if
			RS.Close

		sSQL="Select * from ppoitem where sprodcode='" & sprodcode & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					RS("sprodname")=sprodname
					RS.Update
					RS.MoveNext
				Loop
			End if
			RS.Close

		sSQL="Select * from preceiveitem where sprodcode='" & sprodcode & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					RS("sprodname")=sprodname
					RS.Update
					RS.MoveNext
				Loop
			End if
			RS.Close


response.write "Update Complete"
%>
