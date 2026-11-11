<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<!-- #include file ="Conn.asp" -->

<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,spono,RS2,str
	  ' Create a recordset and open it for adding records
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		spono=Request("spono")

		sSQL="Select * from pproduct"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					str=RS("sProdcode")
					str=replace(str,"""","")
					RS("sprodcode")=str
					RS.Update
					RS.MoveNext
				Loop
			End if
			RS.Close

		sSQL="Select * from ppritem "
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					str=RS("sProdcode")
					str=replace(str,"""","")
					RS("sprodcode")=str
					RS.Update
					RS.MoveNext
				Loop
			End if
			RS.Close


		sSQL="Select * from ppoitem "
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					str=RS("sProdcode")
					str=replace(str,"""","")
					RS("sprodcode")=str
					RS.Update
					RS.MoveNext
				Loop
			End if
			RS.Close

		sSQL="Select * from preceiveitem "
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not RS.EOF then
				Do While Not RS.EOF 
					str=RS("sProdcode")
					str=replace(str,"""","")
					RS("sprodcode")=str
					RS.Update
					RS.MoveNext
				Loop
			End if
			RS.Close

response.write "Update Complete"
%>
