<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->

<%
	' Author: Worakrit   yodthonglert
	' Version: 1.00
%>	

<%
   Dim  Rs,sSQL,RS2,sSQL2
	Dim sScriptPath,sExptype,Color, sExptype1
	Dim sEmpno, sPoscode, sDivcode, sDepcode, sSeccode, sName
	Dim sFlow, NStep, SStep
	Dim MaxStep, Counter
	Dim Division, Department, Section , Position
	Dim DivisionCode, DepartmentCode, SectionCode , PositionCode
	Dim DEM, DDEM, SEM, DEMName, DDEMName, SEMName, SEMStatus, DEMStatus, DDEMStatus
	Dim App01, App02, App03
	Dim sMode
	SEMStatus = TRUE
	DEMStatus = TRUE
	DDEMStatus = TRUE
	Counter=0
	Color="#F8D7D1"
	sEmpno = Session("sEmpno")
	sPoscode = Session("sPoscode")
	sDivcode = Session("sDivcode")
	sDepcode = Session("sDepcode")
	sSeccode = Session("sSeccode")
	sName = Session("sName")
	sExptype = Request("sExptype")
	sExptype1 = Request("sExptype")
	Session("sExptype") = sExptype
	Session("sExptype1") = sExptype1
	DepartmentCode=Mid(sDepcode,1,1)
	SectionCode=Mid(sDepcode,5,1)
	App01 = Request.Form("App01")
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	IF Request("App01") <> ""  then
		Add_Form()
	Else 
		Show_Form()
	End IF		
%>

<%
	Sub Show_Form()

	sSQL = "SELECT * From aEmployee where sDepcode = " & sDepcode & "order by sDepcode"
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	'Response.Write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
%>

<html>
<head>
<title>Create Approve Step</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
</head>

<body bgcolor="#FFFFFF">
<% 
	RS.MoveFirst
	While Not RS.EOF
' Show Department Manager
	IF Mid(sDepcode,1,1) = DepartmentCode  and Mid(sDepcode,5,1) = SectionCode Then
		Department = sDepcode
		Section = sSeccode
		Position = sPoscode
	End IF
 IF RS("sPoscode") = "30" Then
    SEM = RS("sEmpno")
    SEMName = RS("sName")
 Else
    SEMStatus = False
 End IF

 IF RS("sPoscode") = "20" Then
    DEM = RS("sEmpno")
    DEMName = RS("sName")
 Else
    DEMStatus = False
 End IF

 IF RS("sPoscode") = "21" Then
    DDEM = RS("sEmpno")
    DDEMName = RS("sName")
 Else
    DDEMStatus = False
 End IF

	RS.MoveNext
	Wend
	RS.Close
	Counter="0"
Response.write("sSeccode ="&sSeccode)
%> 
 <br>
                  <form method="post" action="<%=sScriptPath%>">
  <table width="75%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Approve 
        Method by Employee Number</b></font></td>
    </tr>
  </table>
  <br>
  <table width="75%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td valign="top" rowspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Approve 
        Step<br>
        By Department</b></font></td>
      <td width="80%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">1.</font> 
        <input type="text" name="App01" size="15" maxlength="12">
      </td>
    </tr>
    <tr> 
      <td width="80%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">2.</font> 
        <input type="text" name="App02" size="15" maxlength="12">
      </td>
    </tr>
    <tr> 
      <td width="80%" height="2"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">3.</font> 
        <input type="text" name="App03" size="15" maxlength="12">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td valign="top" height="2" colspan="2">&nbsp;</td>
    </tr>

    <tr> 
      <td rowspan="2" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Expense 
        Type <br>
        Approver </b></font></td>
    </tr>
    <tr> 
      <td width="80%" height="27"> 
<%
		Show_EmpByExp(sExptype)
%>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td valign="top" height="2" colspan="2"> </td>
    </tr>
    <tr> 
      <td valign="top" rowspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>PUR 
        Staff</b></font></td>
      <td width="80%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">1.</font> 
        <input type="text" name="PUR01" size="15" maxlength="12">
      </td>
    </tr>
    <tr> 
      <td width="80%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">2.</font> 
        <input type="text" name="PUR02" size="15" maxlength="12">
      </td>
    </tr>
    <tr> 
      <td width="80%"> 
		<input type=HIDDEN name=EVENT value="<%=sMode%>">
		<input type=HIDDEN name=sExptype value='<%=Session("sExptype")%>'>
      <input type="submit" name="Submit" value="Submit">
      <input type="submit" name="Reset" value="Reset">
      </td>
    </tr>
  </table>
  </form></body></html>

<% End Sub %>

<%
	Sub Add_Form()
	Dim App01, App02, App03
	Dim Exp01, Exp02, Exp03
	Dim PUR01, PUR02
	Dim sDepcode, DepartmentCode , SectionCode
	Dim bError, FlowStep,  App(5), Counter, Flow, Comm
	Dim Emp(10), nStep
	Dim sFlow, sLevel, sFlowDesc
	FlowStep = 0		
	Counter = 0	
	bError = False

'	Response.write("<BR>sDepcode = "&Session("sDepcode"))
	sDepcode  = Session("sDepcode")
	sPoscode = Session("sPoscode")

		sSQL = "SELECT * FROM pAppflow ORDER BY  sFlow DESC" 
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Rs.Open sSQL, conn, adOpenStatic, adLockOptimistic, adCmdText

		sFlow = RS("sFLow") 
		Response.write("<BR>sFlow = "&sFLow )
	IF Mid(sDepcode,5,1) <> 1  Then
			sFlow = sFlow + 1
			sLevel = "DEM"
			sFlowDesc = "Start at Department Manager"	

			IF sPoscode =  21  Then
				sLevel = "DDEM"
				sFlowDesc = "Start at Deputy Department Manager"	
			End IF

			IF Mid(sDepcode,3,1) <> 1  Then
				sLevel = "DIM"
				sFlowDesc = "Start at Division Manager"	
			End IF
	Else
			sFlow = sFlow + 1 
			sLevel = "SEC"
			sFlowDesc = "Start at Section Manager"
	End IF

		Response.write("<BR>sFlow = "&sFLow )
		Rs.AddNew
		RS("sFlow") = Right("00"&sFlow,3)
		RS("sLevel") = sLevel
		RS("sFlowDesc") = sFlowDesc
		Rs.Update
		Rs.Close

	App01 = Request.Form("App01")
	App02 = Request.Form("App02")
	App03 = Request.Form("App03")
	Exp01 = Request.Form("Exp01")
	Exp02 = Request.Form("Exp02")
	PUR01 = Request.Form("PUR01")
	PUR02 = Request.Form("PUR02")

	IF App01 = "" Then
		bError = True		
	End IF
	
	IF App01 <> "" Then
		IF App02 <> "" Then
			IF App03 <> "" Then
				FlowStep = 7
			Else 
				FlowStep = 6
			End IF
		Else 
			IF App03 <> "" Then
				bError = True
			Else
				FlowStep = 5
			End IF
		End IF
	End IF

	IF bError <> True Then
		sSQL = "SELECT * FROM pAppstep Order by sFlow DESC" 
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText

		IF RS.BOF Then
			sFlow = 1 
		Else 
			sFLow = RS("sFlow")
			sFlow = sFlow + 1 
		End IF
		Session("sFlow") = sFlow
	Response.write("<BR>sFlow = " &sFlow)
	Response.write("<BR>nStep = "&FlowStep)
'		RS.AddNew
'		RS("sFlow") = sFlow 
'		RS("nStep") = FlowStep
'		RS("sStep") 
'		RS("sEmpno")
'		RS("sDESC")
'		RS.Update
		Rs.Close

	FlowStep = FlowStep - 4
	Counter = FlowStep
	Response.write("<BR>Counter = " &Counter)
	While Not Counter = 0
		Comm = "App0"&Counter
		App(Counter) = Request(Comm)
		Response.write("<BR>App("&Counter&") = "&App(Counter))
		Counter = Counter - 1
	Wend

	Response.write("<BR>sFlow = " &sFlow)
	Response.write("<BR>nStep = " &FlowStep+4)
	Counter = FlowStep
' Add Manager Approver 
	While Not Counter = 0
		sSQL = "SELECT * FROM pAppstep"
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Rs.Open sSQL, conn, adOpenStatic, adLockOptimistic, adCmdText
'		RS.Open sSQL , Conn , 1, 3
		RS.AddNew
		RS("sFlow") = Right("00"&sFlow,3)
		RS("nStep") = Counter
		RS("sStep") = "0"&Counter
		RS("sEmpno") =  App(Counter)
		RS("sDESC") = ""
'		RS("sFlow") = "001"
'		RS("nStep") = Counter
'		RS("sStep") = "01"
'		RS("sEmpno") =  App(Counter)   
		RS.Update
		Counter = Counter - 1 
	Wend
' End Manager Approver

' Add Exptype Approver 
		FlowStep = FlowStep + 1
		RS.AddNew
		RS("sFlow") = Right("00"&sFlow,3)
		RS("nStep") = FlowStep
		RS("sStep") = "0"&FlowStep
		RS("sEmpno") =  Exp01
		RS("sDESC") = ""
		RS.Update

		FlowStep = FlowStep + 1
		RS.AddNew
		RS("sFlow") = Right("00"&sFlow,3)
		RS("nStep") = FlowStep
		RS("sStep") = "0"&FlowStep
		RS("sEmpno") =  Exp02
		RS("sDESC") = ""
		RS.Update
' End Exptype Approver

' Add PUR Approver 

		FlowStep = FlowStep + 1
		RS.AddNew
		RS("sFlow") = Right("00"&sFlow,3)
		RS("nStep") = FlowStep
		RS("sStep") = "0"&FlowStep
		RS("sEmpno") =  PUR01
		RS("sDESC") = ""
		RS.Update

		FlowStep = FlowStep + 1
		RS.AddNew
		RS("sFlow") = Right("00"&sFlow,3)
		RS("nStep") = FlowStep
		RS("sStep") = "0"&FlowStep
		RS("sEmpno") =  PUR02
		RS("sDESC") = ""
		RS.Update
		RS.Close
' End PUR Approver
		sFlow = Right("00"&sFlow,3)

		Counter = FlowStep	
	sSQL = "SELECT * FROM pAppstep where sFlow = " &sFlow& " order by nStep"
	sSQL2 = "SELECT * FROM pPrapp"
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Rs.Open sSQL, conn, adOpenStatic, adLockOptimistic, adCmdText
	Rs2.Open sSQL2, conn, adOpenStatic, adLockOptimistic, adCmdText

	While Not  Counter = 0
		While Not RS.EOF
			nStep = RS("nStep")
					Emp(Counter) = RS("sEmpno")
					Response.write("<BR>nStep = "&RS("nStep"))
					Response.write("<BR>Empno = "&RS("sEmpno"))
					Response.write("<BR>sPrno = "&Session("sPrno"))
					Response.write("<BR>Counter = "&Counter)
			Rs2.ADDNew
			Rs2("sEmpno")  =  RS("sEmpno")
			Rs2("sPrno") = Session("sPrno")
			Rs2("nStep") = "0"&FlowStep
			Rs2("sStep") = "0"&nStep
			Rs2.Update
			RS.MoveNext
		Wend	
	Counter = Counter - 1 
	Wend

	Rs2("DREQDATE") = Day(date)
	Rs2("TREQDATE") = FormatDateTime(Now,4)
	Rs2.Update
	RS.Close
	RS2.Close	
' Add Approve Step 2 Table > pPrapp

	Session("sFlow") = ""
	Session("sExptype") = ""
	Else
		Show_Form()
	End IF
%>

<% End Sub %>

<%
	Sub Show_EmpByExp(sExptype)
'	Dim sSQL2,RS2,sExptype,Counter
%>
		<table width='100%' border='0' cellspacing='0' cellpadding='0'>
<%
	Counter = 0
	sSQL2 = "select * from pExpstep where sExptype = "& sExptype &"order by sStep ASC"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText		
	While Not RS2.EOF
	Counter = Counter + 1
%>
           <tr>
			<td>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="2"> <%=Counter%></font> 
      <input type="text" name='Exp0<%=Counter%>' size="15" maxlength="12" Value='<%=RS2("sEmpno")%>' ReadOnly>
			</td>
          </tr>
<%
	RS2.MoveNext
	Wend 
	Rs2.Close
	Counter = 0
%>
        </table>
<%
Response.write("<BR> sExpyte Sub = " &sExptype)
End Sub
%> 