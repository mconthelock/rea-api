<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../List.asp" -->
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
Dim sEmpnoapp,sEmpno
		sEmpnoapp = Request("sEmpno")
		sEmpno=Request("sEmpno2")
  Set Rs = Server.CreateObject("ADODB.RecordSet")
		  sSQL = "SELECT * FROM  pprapp WHERE sEmpno='" & sEmpno &"'" &_
			" And cApprov='0'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
If Not Rs.EOF then
	Do While Not Rs.EOF
		Rs("sEmpnoapp") = sEmpnoapp
	  Rs.Update	
	Rs.MoveNext
	Loop
End if
	  Rs.Close
		sSQL="Select * from pAppstep where sEmpno='" & sEmpno & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
If Not Rs.EOF then
		Do While Not Rs.EOF 
			RS("sEmpapp")=sEmpnoapp
			RS.Update
		Rs.MoveNext
		Loop
End if
			Rs.Close
	  Set Rs = Nothing
response.redirect "../security/security.asp"
%>