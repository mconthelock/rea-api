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
Dim sEmpnoapp
		sEmpnoapp = Ucase(Request("sEmpno"))

  Set Rs = Server.CreateObject("ADODB.RecordSet")
		  sSQL = "SELECT * FROM  pprapp WHERE sEmpno='" & Session("sEmpno") &"'" &_
			" And cApprov='0'"
'	sSQL = "SELECT * FROM  pprapp WHERE sEmpno='" & Session("sEmpno") &"'"
	Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	Do while Not RS.EOF
		Rs("sEmpnoapp") = sEmpnoapp
		Rs.Update	
		RS.MoveNext
	Loop
	Rs.Close

	sSQL="Select * from pAppstep where sEmpno='" & Session("sEmpno") & "'"
	Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	Do while Not Rs.EOF
		RS("sEmpapp")=sEmpnoapp
		RS.Update
		RS.MoveNext
	Loop
	Rs.Close

	sSQL="Select * from pExpstep where sEmpno='" & Session("sEmpno") & "'"
	Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	Do while Not Rs.EOF
		RS("sEmpapp")=sEmpnoapp
		RS.Update
	RS.MoveNext
	Loop
	Rs.Close
	Set Rs = Nothing
response.redirect "../workflow/workflow.asp"
%>