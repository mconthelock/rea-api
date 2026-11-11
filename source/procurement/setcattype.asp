<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<!-- #include file ="Conn.asp" -->
<!-- #include file="fcommon.asp" -->
<!-- #include file="security/secure.asp" -->
<%
	
    Dim  Rs,sSQL
	Dim sScriptPath,spono,Color
	spono=Request("spono")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
		sSQL="Select * from pproduct"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
if Not RS.EOF then
	Do While Not RS.EOF
		RS("sCattype")=mid(RS("sCatID"),1,3) 
		Rs.Update
		RS.MoveNext
	Loop
		Rs.Close
		Set Rs=Nothing
		response.write "Set Complete"
else
response.write "No Complete"
End if

%>