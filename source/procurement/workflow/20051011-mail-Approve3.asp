<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo
	' Version: 1.00
%>	

<%
	Dim sErrImage, Rs, sSQL,RS7,RS2,rs3,Rs5,sSQL5
	Dim sPrno, nItem, Counter
	Dim	cStatus(100)
	Dim nStep,MaxStep,NextStep
	Dim tAppTime,dAppDate, cApprov,cPrtype,LastStep,chkitem,chkRecMail,sRecmail,sEmpnoRec,mail

		sPrno = Request("sPrno")
		nItem = Request("nItem")
		Counter = Request("Count_Item")



	' Up Nstep One Level

		Set Rs = Server.CreateObject("ADODB.RecordSet")
		sSQL = "SELECT * FROM  ppr WHERE sPrno= '" & sPrno & "'"
		Rs.Open sSQL, conn, 1,3
		If Not Rs.EOF then
			cPrtype = RS("cPrtype")
		End if
		Rs.Close

		sSQL = "SELECT * FROM  pprapp WHERE cApprov='0'"  &_
	" and sPrno= '" & sPrno &_
	"' Order By nStep"
		Rs.Open sSQL, conn, 1,3
			Rs.MoveFirst
			LastStep=RS("nStep")				
			Rs.Close

if Cint(Laststep) <30 then
	sSQL = "Select * From pPrapp where (sEmpno = '"&Session("sEmpno")&"' or sEmpnoapp='" & Session("sEmpno") & "') and sPrno ='"&sPrno&"'" &_
	" and capprov='0'" &_
	" and nstep <'30'" &_
	" Order By nStep"
	Rs.Open sSQL, Conn, 1,3
		If Not Rs.EOF then
			Do While Not RS.EOF
				RS("DAPPDATE")= Date
				tApptime=hour(now) & ":" & minute(now)
				Rs("tApptime")= tApptime
				RS("cApprov") = "1"
				RS("sComment")=Session("sEmpno")
				RS.Update
			RS.MoveNext
			Loop
		End if
		Rs.Close
else
	sSQL = "Select * From pPrapp where (sEmpno = '"&Session("sEmpno")&"' or sEmpnoapp='" & Session("sEmpno") & "') and sPrno ='"&sPrno&"'" &_
	" and capprov='0'" &_
	" and nstep <>'35'" &_
	" Order By nStep"
	Rs.Open sSQL, Conn, 1,3
		If Not Rs.EOF then
			Do While Not RS.EOF
				RS("DAPPDATE")= Date
				tApptime=hour(now) & ":" & minute(now)
				Rs("tApptime")= tApptime
				RS("cApprov") = "1"
				RS("sComment")=Session("sEmpno")
				RS.Update
			RS.MoveNext
			Loop
		End if
		Rs.Close
end if

chkPritem()

'Urgent and Not reject

if cPrtype="1" and chkitem<>"0" then
	sSQL = "Select * From pPrapp where sPrno ='"&sPrno&"'" &_
	" and capprov='0'" &_
	" and nStep < '30'"
	Rs.Open sSQL, Conn, 1,3
	if Not RS.EOF Then
		Do While Not RS.EOF 
			RS("DAPPDATE")= Date
			tApptime=hour(now) & ":" & minute(now)
			Rs("tApptime")= "-"
			If Mid(RS("nStep"),1,1) <> "3" then
				RS("cApprov") = "1"
				RS("sComment")=RS("sEmpno")
			Else
				RS("cApprov") = "0"
			End if	
'			RS("sComment")=Session("sEmpno")
			RS.Update
			Rs.MoveNext
		Loop
	End if
	Rs.Close
End if
if cPrtype="1" and cint(Laststep)=30 then

		Set Rs3 = Server.CreateObject("ADODB.RecordSet")
		sSQL = "SELECT * FROM  ppritem WHERE sPrno= '" & sPrno & "'" &_
		" and (cStatus='1'" &_
		" or cStatus='2'" &_
		" or cStatus='9')"
		Rs3.Open sSQL, conn, 1,3
		If Rs3.EOF then
		   NextStep="0"
		else
 		  NextStep="1"
		End if
		Rs3.Close

	IF NextStep="0" then 
				sSQL = "Select * From pPrapp where sPrno ='"&sPrno&"'" &_
				" and capprov='0'" &_
				" and nStep = '30'" 
				Rs.Open sSQL, Conn, 1,3
				if Not RS.EOF Then
					Do While Not RS.EOF 
						RS("DAPPDATE")= Date
						tApptime=hour(now) & ":" & minute(now)
						Rs("tApptime")= "-"
							RS("cApprov") = "1"
							RS("sComment")=RS("sEmpno")
						RS.Update
						Rs.MoveNext
					Loop
				End if
				Rs.Close
				sSQL = "SELECT * FROM  ppr WHERE sPrno= '" & sPrno & "'"
				Rs.Open sSQL, conn, 1,3
				RS("nStep") = cInt(NextStep)
				RS("dFinishdate")=date
				RS.Update
				Rs.Close
				Response.Redirect "workflow.asp"
	Else
				sSQL = "Select * From pPrapp where sPrno ='"&sPrno&"'" &_
				" and capprov='0'" &_
				" and nStep > '30'" &_
				" and nStep < '35'"
				Rs.Open sSQL, Conn, 1,3
				if Not RS.EOF Then
					Do While Not RS.EOF 
						RS("DAPPDATE")= Date
						tApptime=hour(now) & ":" & minute(now)
						Rs("tApptime")= "-"
							RS("cApprov") = "1"
							RS("sComment")=RS("sEmpno")
						RS.Update
						Rs.MoveNext
					Loop
				End if
				Rs.Close
	End IF
		sSQL = "SELECT * FROM  ppr WHERE sPrno= '" & sPrno & "'"
		Rs.Open sSQL, conn, 1,3
			nStep = cInt(RS("nStep"))
				RS("nStep") = 35
			RS.Update
		Rs.Close
				Set RS7=Server.CreateObject("ADODB.Recordset")
					sSQL="Select * from ppr Where sprno='" & sPrno & "'"
			 	Rs7.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
				If Not RS7.EOF then
					Do While Not Rs7.EOF
							RS7("cApprove") = "1"
							RS7("dFinishdate")=date
							RS7.Update
							Rs7.MoveNext
					Loop
				End if
			RS7.Close
			

			
			Set Rs7=nothing
			Response.redirect  "../workflow/createpo.asp?sprno=" & sPrno

End if

		sSQL = "SELECT * FROM  pprapp WHERE cApprov='0'"  &_
	" and sPrno= '" & sPrno &_
	"' Order By nStep"
		Rs.Open sSQL, conn, 1,3
			Rs.MoveFirst
			NextStep=Rs("nStep")
			sEmpnoRec=RS("sEmpno")
			Rs.Close

		chkReceiveMail()
		IF chkRecMail="Y" then
			Set Mail=Server.CreateObject("CDONTS.NewMail")
			Mail.From="adminprocurement@amec.ina.melco.co.jp"
			Mail.To=RecMail(sEmpnoRec)'"eak_liverpool@hotmail.com"
			Mail.Subject=sprno
			Mail.Body= sprno & " Waiting for your approve"
			Mail.Send
			Set Mail=nothing
		End IF


		Set Rs3 = Server.CreateObject("ADODB.RecordSet")
		sSQL = "SELECT * FROM  ppritem WHERE sPrno= '" & sPrno & "'" &_
		" and (cStatus='1'" &_
		" or cStatus='2'" &_
		" or cStatus='9')"
		Rs3.Open sSQL, conn, 1,3
		If Rs3.EOF then
		   NextStep=0
		End if
		Rs3.Close

		sSQL = "SELECT * FROM  ppr WHERE sPrno= '" & sPrno & "'"
		Rs.Open sSQL, conn, 1,3
			nStep = cInt(RS("nStep"))
			if cPrtype="1" and LastStep="10" and cDbl(Nextstep)<>0 then
				RS("nStep") = 30
			else
				RS("nStep") = NextStep
			End if
			RS("dFinishdate")=date
			RS.Update
		Rs.Close


		Set Rs = Nothing					

		Set Rs = Server.CreateObject("ADODB.RecordSet")
		sSQL = "SELECT * FROM  pprapp WHERE sPrno= '" & sPrno & "'" &_
		" Order By nStep desc"
		Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF Not Rs.EOF then
				Rs.MoveFirst
				RS.MoveNext
				maxstep=RS("nStep")
				'response.write maxstep
				'response.write nstep
		End if
		If Cint(MaxStep)=Cint(nStep) then
				Set RS7=Server.CreateObject("ADODB.Recordset")
					sSQL="Select * from ppr Where sprno='" & sPrno & "'"
			 	Rs7.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
				If Not RS7.EOF then
					Do While Not Rs7.EOF
							RS7("cApprove") = "1"
							RS7("dFinishdate")=date
							RS7.Update
							Rs7.MoveNext
					Loop
				End if
							RS7.Close
			Response.redirect  "../workflow/createpo.asp?sprno=" & sPrno
		Else
		Set Rs5 = Server.CreateObject("ADODB.RecordSet")
		sSQL5="Select * from pprapp where sPrno='" & sPrno & "'" &_
		" and cApprov='0'" &_
		" Order By nStep"
		Rs5.Open sSQL5, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS5.EOF then
			RS5.MoveFirst
			if Cint(RS5("nStep"))=30 then
				Move()
			End if
		End if
		RS5.Close
		Set Rs5=Nothing
			Response.Redirect "workflow.asp"
		End if
%>
<%
Sub Move()
Dim sSQL2,RS2,Rs3
		sSQL2= "Select * from ppritem Where sPrno='" & sprno & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Set Rs3 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		Rs3.Open "ppritem2", conn, adOpenKeySet, adLockPessimistic, adCmdTable
	IF NOT Rs2.EOF  Then
		Do while Not RS2.EOF
		  Rs3.AddNew
			Rs3("sPrno") = RS2("sPrno")
			Rs3("nItem") = RS2("nItem")
			Rs3("sProdcode") = RS2("sProdcode")
			Rs3("sProdname") = RS2("sProdname")
			Rs3("sVendcode") =RS2("sVendcode")
			Rs3("sProddesc") =RS2("sProddesc")
			Rs3("nQuality") =RS2("nQuality")
			Rs3("nPrice") =RS2("nPrice")
			Rs3("nAmount") =RS2("nAmount")
			Rs3("sCurcode") =Rs2("sCurcode")
			Rs3("sUseFor") =RS2("sUseFor")
			Rs3("sQuatation") =RS2("sQuatation")
			Rs3("sUnitcode") =RS2("sUnitcode")
			Rs3("sComment") =RS2("sComment")        
			Rs3("sNote") =RS2("sNote")        
			Rs3("cStatus") =RS2("cStatus")        
			Rs3("sImagespec") =Rs2("sImagespec")        
		  Rs3.Update	
		Rs2.MoveNext
		Loop
	End IF
	RS2.Close
	Set RS2=Nothing
'	RS3.Close
'	Set RS3=Nothing
End Sub
%>
<%
Sub chkPritem()
Dim Rs10,sSQL10
	Set Rs10 = Server.CreateObject("ADODB.RecordSet")
	sSQL10="Select * from ppritem where sprno='" & sprno & "'" &_
	" and cstatus<>'0'" 
	Rs10.Open sSQL10, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if RS10.EOF then
		chkitem="0"
	else
		chkitem="1"
	End if
	RS10.Close
	Set RS10=Nothing
End Sub

Sub chkReceiveMail()
Dim Rs10,sSQL10
	Set Rs10 = Server.CreateObject("ADODB.RecordSet")
	sSQL10="Select  cRecMail from  aEmployee where sEmpno='" & sEmpnoRec & "'"
	Rs10.Open sSQL10, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if  Not RS10.EOF then
		chkRecMail=RS10("cRecMail")
	End if
	RS10.Close
	Set RS10=Nothing
End Sub

Function RecMail(sEmp)
Dim Rs10,sSQL10
	Set Rs10 = Server.CreateObject("ADODB.RecordSet")
	sSQL10="Select  sRecMail from  aEmployee where sEmpno='" & sEmp & "'"
	Rs10.Open sSQL10, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if  Not RS10.EOF then
		sRecMail=RS10("sRecmail")
	End if
	RS10.Close
	Set RS10=Nothing
	RecMail=sRecMail
End Function
%>