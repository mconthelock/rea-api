<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<%
Dim RS,sSQL,sSQL2,RS2,RS3,RS4,RS5,RS6,RS7,RS8,RsG, RsCheck
Dim sRequester,sDivcode,sDepcode,sSeccode,sFlow,str,sStep,nStep,sEmpno,sPrno,dreqdate,chkRecMail
Dim treqtime,sExptype,FirstnStep,MinStep,cPrtype,sEmpapp,sPoscode,Firststep,Mail,sEmpnoRec,bdgNo,biddingno
biddingno=Request("biddingno")
sRequester=request("srequester")
sPrno=Request("sPrno")
dReqdate=Request("dreqdate")
  bdgNo=Request("bdgNo")
  'response.write dbgNo
  'response.write "tttttttt"&bdgNo
  'response.write dReqdate
'response.end
sExptype=Request("sExptype")
cPrtype=Request("cPrtype")
Set Rs=Server.CreateObject("ADODB.Recordset")
sSQL="Select * from aEmployee Where sEmpno='" & sRequester & "'"
Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
IF Not RS.EOF  then
	Do While Not Rs.EOF
		sDivcode=RS("sDivcode")
		sDepcode=RS("sDepcode")
		sSeccode=RS("sSeccode")
		sPoscode=RS("sPoscode")
	RS.MoveNext
	Loop

if sDivcode<>"00" and sDepcode<>"00" and sSeccode <> "00" then
	str= sDivcode & "," & sDepcode & "," & sSeccode
elseif sDivcode<>"00" and sDepcode="00" and sSeccode = "00" then
	str=sDivcode
elseif sDivcode<>"00" and sDepcode<>"00" and sSeccode = "00" then
	str= sDivcode & "," & sDepcode
end if


'By Supamid on 23/06/2023 insert data for packing to table PPRPACKAGE
'By Supamid on 27/03/2026 Check SPRNO exist Before insert 
	if trim(biddingno) <> "" then
		' เช็คก่อนว่ามี SPRNO นี้หรือยัง
		Set RsCheck = Server.CreateObject("ADODB.RecordSet")
		RsCheck.Open "SELECT 1 FROM PPRBIDDING WHERE SPRNO = '" & sPrno & "'", conn, adOpenStatic, adLockReadOnly, adCmdText

		if RsCheck.EOF then
			' ไม่มี -> ค่อย insert
			Set RsG = Server.CreateObject("ADODB.RecordSet")
			RsG.Open "PPRBIDDING", conn, adOpenKeySet, adLockPessimistic, adCmdTable

			RsG.AddNew
				RsG("SPRNO") = sPrno
				RsG("BIDDINGNO") = biddingno
			RsG.Update    

			RsG.Close
			Set RsG = Nothing
		end if

		RsCheck.Close
		Set RsCheck = Nothing

	end if

	
	
Set Rs2=Server.CreateObject("ADODB.Recordset")

sSQL="Select * from vpappstep Where sSteptype=1 and slevel='" & str & "'" &_
" and sEmpno='" & sRequester & "'" &_
" Order By nstep"
Rs2.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
If Not RS2.EOF then
	Firststep=RS2("sStep")
End if
Rs2.Close

sSQL="Select * from vpappstep Where sSteptype=1 and slevel='" & str & "'" &_
" Order By nstep"
'response.write sSQL
'response.end
Rs2.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText

Set Rs3=Server.CreateObject("ADODB.Recordset")
Rs3.Open "pprapp", conn, adOpenKeySet, adLockPessimistic, adCmdTable
if Not RS2.EOF then
	Do While Not Rs2.EOF
		nStep=RS2("nStep")
		sStep=RS2("sStep")
		sEmpno=RS2("sEmpno")
		sEmpapp=RS2("sEmpapp")
	
		if sStep="01"  then 
			RS3.AddNew
			RS3("sEmpno")= sRequester'sEmpno
			RS3("sPrno")=sPrno
			''RS3("dReqdate")=dReqdate
			RS3("dReqdate")=date
			treqtime=hour(now) & ":" & minute(now)
			Rs3("tReqtime")= treqtime
			Rs3("nStep")=nStep
			Rs3("sStep")=sStep
			RS3("dAppdate")=date
			Rs3("tApptime")=hour(now) & ":" & minute(now)
			RS3("cApprov")="5"
			Rs3("sComment")=sRequester
			RS3.Update
		Else
			If sRequester<>sEmpno and cDbl(Firststep) < cDbl(sstep) then
				RS3.AddNew
				RS3("sEmpno")= sEmpno
				RS3("sPrno")=sPrno
				''RS3("dReqdate")=dReqdate
				RS3("dReqdate")=date
				treqtime=hour(now) & ":" & minute(now)
				Rs3("tReqtime")= treqtime
				Rs3("nStep")=nStep
				Rs3("sStep")=sStep
				if cprtype="1" and (sPoscode="20" Or sPoscode="30") then
					RS3("cApprov")="1"
					RS3("sComment")=sEmpno
					RS3("dappdate")=date
					RS3("tApptime")="-"
				else
					RS3("cApprov")="0"
				end if
					RS3("sEmpnoapp")=sEmpapp
					RS3.Update
			End if
		End if

		RS2.Movenext
	Loop
	RS2.Close
End if

'---> Expense type Approve

		Set Rs4=Server.CreateObject("ADODB.Recordset")
		sSQL="Select * from pexpstep where sexptype='" & sExptype & "'" &_
		" Order By nStep"
		RS4.Open sSQL,conn,AdOpenKeySet,adLockPessimistic,adCmdText

		IF Not Rs4.EOF then
			Do While Not RS4.EOF
				nStep=RS4("nStep")
				sStep=RS4("sStep")
				sEmpno=RS4("sEmpno")
				sEmpapp=RS4("sEmpapp")
				If sRequester<>sEmpno then
					RS3.AddNew
					RS3("sEmpno")=sEmpno
					RS3("sPrno")=sPrno
					''RS3("dReqdate")=dReqdate
					RS3("dReqdate")=date
					treqtime=hour(now) & ":" & minute(now)
					Rs3("tReqtime")= treqtime
					Rs3("nStep")=nStep
					Rs3("sStep")=sStep
						if cprtype="1" and (sPoscode="20" Or sPoscode="30" or sPoscode <=10) then
							RS3("cApprov")="1"
							RS3("sComment")=sEmpno
							RS3("dappdate")=date
							RS3("tApptime")="-"
						else
							RS3("cApprov")="0"
						end if
					RS3("sEmpnoapp")=sEmpapp
					RS3.Update
				End if
			RS4.MoveNext
			Loop
		End if


Set Rs5=Server.CreateObject("ADODB.Recordset")
sSQL="Select * from vpappstep Where sSteptype=3 and slevel='" & str & "'" &_
" Order By nstep"
Rs5.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
	if Not RS5.EOF then
		Do While Not Rs5.EOF
			nStep=RS5("nStep")
			sStep=RS5("sStep")
			sEmpno=RS5("sEmpno")
			sEmpapp=RS5("sEmpapp")
'			If sRequester<>sEmpno then
				RS3.AddNew
				RS3("sEmpno")=sEmpno
				RS3("sPrno")=sPrno
				''RS3("dReqdate")=dReqdate
				RS3("dReqdate")=date
				treqtime=hour(now) & ":" & minute(now)
				Rs3("tReqtime")= treqtime
				Rs3("nStep")=nStep
				Rs3("sStep")=sStep
				RS3("cApprov")="0"
				RS3("sEmpnoapp")=sEmpapp
				RS3.Update
'			End if
		RS5.Movenext
		Loop
		RS5.Close
	End if

Set Rs7=Server.CreateObject("ADODB.Recordset")
sSQL="Select * from pprapp where sprno='" & sprno & "'" &_
" and cApprov='0'" &_
" Order By nStep"
Rs7.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
if Not Rs7.EOF then
	Rs7.MoveFirst
'	Rs7.MoveNext	
	MinStep= RS7("nStep")
	sEmpnoRec=RS7("sEmpno")
End if

Set Rs8=Server.CreateObject("ADODB.Recordset")
sSQL="Select  * from ppr where sprno='" & sprno & "'"
Rs8.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
If Not RS8.EOF then
	Do While Not Rs8.EOF
'if cPrtype="1" and sSeccode="00" then
'		Rs8("nStep")=30
'else
		Rs8("nStep")=MinStep
'End if
	RS8.MoveNext
	Loop
End if
Rs7.Close

	sSQL="Select * from ppr where sprno='" & sprno & "'"
	Rs7.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
	RS7("cApprove")="7"
	RS7.Update
	RS7.Close	
	chkReceiveMail()

	IF chkRecMail="Y" then
'		Set Mail=Server.CreateObject("CDONTS.NewMail")
''			Set Mail=Server.CreateObject("CDO.Message")
'			Mail.From="adminprocurement@amec.ina.melco.co.jp"
''			Mail.From="adminprocurement@MitsubishiElevatorAsia.co.th"
''			Mail.To=RecMail(sEmpnoRec)'"eak_liverpool@hotmail.com"
''			Mail.Subject=sprno
'			Mail.Body= sprno & "Waiting for your approve"
''			Mail.TextBody = sprno & " Waiting for your approve"
''			Mail.Send
''		Set Mail=nothing
		sendMail "amecoutlook", "adminprocurement", "adminprocurement@MitsubishiElevatorAsia.co.th", RecMail(sEmpnoRec), sprno, sprno & " Waiting for your approve"
		'sendMail "amec1", "adminprocurement", "adminprocurement@MitsubishiElevatorAsia.co.th", "kanittha@MitsubishiElevatorAsia.co.th", sprno, sprno & " Waiting for your approve"
	End IF
if Request("go")="next" then
	response.redirect "../workflow/addpr.asp?sExptype=" & Request("sExptype") 
else
	response.redirect "../workflow/workflow.asp"
end if
End if

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
Dim Rs10,sSQL10,sRecMail
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