<%

Sub List_termcode(sterm)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from ptermcode Order By stermcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")

		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("stermcode") = sterm Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("stermcode")& "'" & cSelect &">"& Rs2("stermdesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

Sub List_Pexptype(sexptype)
Dim sSQL2,RS2,cSelect
	sSQL2= "Select * from pexpenstype where SEXPSTATUS = '1' Order By sExptype"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("sexptype") = sexptype Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			IF Rs2("sexptype") = "54" and Session("sDivcode") <> "040101" then 
				Rs2.MoveNext
			END IF		
			IF Rs2("sexptype") = "B0" and Session("sDepcode") <> "060801" then 
				Rs2.MoveNext
			END IF				
			Response.Write "<option value= '"&Rs2("sexptype")& "'" & cSelect &">"&Rs2("sexptype")&"-" & Rs2("sexpdesc") &"</option>"& vbCrLf
			'Response.Write "<option value= '"&Rs2("sexptype")& "'" & cSelect &">"&Rs2("sexptype")&"-" & Rs2("sexpdesc") 
			'IF Rs2("sexptype") = "54" and Session("sDivcode") <> "30101" then 
			'	Response.Write "<option value= '"&Rs2("sexptype")& "'" & cSelect &">"&Rs2("sexptype")&"-" & Rs2("sexpdesc") &"</option>"& vbCrLf
			'else 
			'	Rs2.MoveNext
			'END IF
			'Response.Write "</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub



Sub pExptype_Desc(sexptype)
Dim sSQL2,RS2, sExpDESC
		sSQL2= "Select * from pExpenstype"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		While NOT Rs2.EOF
			IF RS2("sExptype") = sExptype Then 
				sExpDESC = RS2("sExpDESC")	
				Response.write sExpDESC
    		End IF
		Rs2.MoveNext
	 	Wend
	End IF
End Sub

Sub Show_Pexptype(sexptype)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pexpenstype Where sexptype='" & sexptype & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sexptype") = sexptype Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write(Rs2("sexpdesc"))
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Pvender(svendcode)
Dim sSQL2,RS2,cSelect
		sSQL2="Select * from pvender Order By svendcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("svendcode") = svendcode Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("svendcode")& "'" & cSelect &">"& Rs2("stname") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub Show_Vender(Vendcode)
Dim sSQL2,RS2,cSelect
		sSQL2="Select * from pvender where sVendcode='" & Trim(Vendcode) & "' " &_
		"Order By svendcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	IF NOT Rs2.EOF  Then
		Response.Write(Rs2("sEname"))
	End IF
End Sub

Sub List_Ppr(sprno)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from Ppr"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sprno") = sprno Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sprno")& "'" & cSelect &">"& Rs2("sprno") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Pflowstep(sStep)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from Pflowstep order by sstep"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sStep") = sStep Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sStep")& "'" & cSelect &">"& Rs2("sStep") &" - "&Rs2("sStepDesc")&"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Pappflow(sFlow)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from Pappflow order by sflow"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sFlow") = sFlow Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sFlow")& "'" & cSelect &">"& Rs2("sFlow") &" - "&Rs2("sFlowDesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Program(sProgcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from aProgram Order By sProgcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sProgcode") = sProgcode  Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sProgcode")& "'" & cSelect &">"& Rs2("sProgdesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub Show_Program(sProgcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from aProgram where sProgcode='" & sProgcode & "'" &_
		"Order By sProgcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Response.write(RS("sProgdesc"))
	End if
End Sub

Sub Count_Item(sPrno)
	Dim sSQL2,RS2,cSelect,Count_Item
		sSQL2= "Select Count(nItem) as Item from ppritem where sprno='" & sPrno & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Response.write(RS2("Item"))
	End if	
End Sub

Function GetMonthName(m)
	Select Case m
		Case 1
			GetMonthName = "Jan"
		Case 2
			GetMonthName = "Feb"
		Case 3
			GetMonthName = "Mar"
		Case 4
			GetMonthName = "Apr"
		Case 5
			GetMonthName = "May"
		Case 6
			GetMonthName = "Jun"
		Case 7
			GetMonthName = "Jul"
		Case 8
			GetMonthName = "Aug"
		Case 9
			GetMonthName = "Sep"
		Case 10
			GetMonthName = "Oct"
		Case 11
			GetMonthName = "Nov"
		Case 12
			GetMonthName = "Dec"
	End Select
End Function

Function MonthDay(m, y)
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12
            MonthDay = 31
        Case 2
            If y Mod 4 = 0 Then
                If y Mod 100 = 0 And y Mod 1000 <> 0 Then
                    MonthDay = 28
                Else
                    MonthDay = 29
                End If
            Else
                MonthDay = 28
            End If
        Case 4, 6, 9, 11
            MonthDay = 30
    End Select
End Function

Function GetDayName(d)
	Select Case d
		Case 1
			GetDayName = "Sun"
		Case 2
			GetDayName = "Mon"
		Case 3
			GetDayName = "Tue"
		Case 4
			GetDayName = "Wen"
		Case 5
			GetDayName = "Thu"
		Case 6
			GetDayName = "Fri"
		Case 7
			GetDayName = "Sat"
	End Select
End Function

Function Show_Date(dDate)
Dim dday,dmonth,dyear
	dday=Day(dDAte)
	dmonth=GetMonthName(Month(dDate))
''	dYear=Year(dDate)
'' Modify by Kallayanee on 29/08/2008
	dYear=Year(dDate)+543
'	Show_Date=" " & dday & " " & dmonth & " " & dYear+543
	Show_Date=" " & dday & " " & dmonth & " " & dYear
End Function

Sub List_Currency(sCurcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from Pcurrency Order By sCurrency"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("scurcode") = sCurcode Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
		Response.Write "<option value= '" & Rs2("sCurcode") & "'" & cSelect &">"& Rs2("sCurrency") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Function Show_Pos(sempno)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from aemployee Where sempno='" & sempno & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Show_pos=(Rs2("sposcode"))
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set RS2=Nothing
End Function

Function Show_CatalogDES(sCattype)
Dim sSQL2,RS2,cSelect
		sSQL2= "select * from pCattype Where sCattype='" & sCattype & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF  Then
				Response.write Rs2("SCATDESC")
		End IF
End Function

Function ShowCattype(sCattype)
Dim sSQL2,RS2,cSelect
		sSQL2= "select * from pCattype Where sCattype='" & sCattype & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF  Then
				Response.write Rs2("SCATDESC")
		End IF
End Function

Function ShowCatalog(sCatID)
Dim sSQL2,RS2,cSelect
		sSQL2= "select * from pCatalog Where sCatID='" & sCatID & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF  Then
				Response.write Rs2("SCATALOG")
		End IF
End Function

Function Show_Product(sProdcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "select * from pProduct Where sProdcode='" & sProdcode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF  Then
				Response.write Rs2("STPRODNAME")
		End IF
End Function


Function Show_Div(sempno)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from aemployee Where sempno='" & sempno & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Show_Div=mid((Rs2("sdivcode")),1,3)
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set RS2=Nothing
End Function

Sub List_Division(Divcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pDivision Where sDivcode='" & Divcode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Response.write Rs2("sDivision")
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Department(DepCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pDepartment Where sDepcode='" & Depcode & "'" &_
		" Order By sDepartment"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Response.write Rs2("sDepartment")
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Sec(SecCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pSection Order By sSection"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
					If Rs2("sSeccode") = seccode Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sSeccode")& "'" & cSelect &">"& Rs2("sSec") &"</option>"& vbCrLf

			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Section(SecCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pSection Where sSeccode='" & Seccode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF  Then
				Response.write Rs2("sSection")
		End IF
End Sub

Sub Desc_Section(SecCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pSection Where sSeccode='" & Seccode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF  Then
				Response.write Rs2("sSection")
		End IF
End Sub

Sub List_Position(PosCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pPosition Where sPoscode='" & Poscode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Response.write Rs2("sPosition")
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Status(StatusCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pStatusEmp Where sStatuscode='" & Statuscode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Response.write Rs2("sStatusDesc")
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub List_Unit(sUnitcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from PUnit WHERE CSTATUS = '1' Order By sUnitdesc"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sUnitcode") = sUnitcode Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
		Response.Write "<option value= '" & Rs2("sUnitcode") & "'" & cSelect &">"& Rs2("sUnitDesc") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub

Sub Show_Unit(sUnitcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from PUnit where sUnitcode='" & sUnitcode & "'" &_
		" Order By sUnitcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Response.Write Rs2("sUnitDesc")
	End IF
End Sub

Sub Show_Currency(sCurcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from Pcurrency Where sCurcode = '"&sCurcode&"'"
'		Response.write sSQL2
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Response.Write Rs2("sCurrency")
	End IF
End Sub

Sub Show_Empno(sEmpno)
    Dim  Rs,sSQL,sStatus,sScriptPath,smode
	Dim scurcode,scurrency
	Dim img_sEmpno, img_sName, img_sSurname, img_sTName,  img_sTSurname, img_sLogin ,img_sPassword
	Dim img_sDivision, img_sDepartment, img_sSection, img_sPosition
	Dim sDivision, sDepartment, sSection, sPosition
	Dim sDivcode, sDepcode, sSeccode, sPoscode
	Dim sDivc, sDepc, sSecc, sPO, sName, sSurname, sTName, sTSurname
	Dim sLogin, sPassword, sPassword1
	Dim bError, sbread
	Dim sErrImage
	bError = False
	sErrImage = "<IMG SRC='../images/icerror.gif'>"
sScriptPath = Request.ServerVariables("SCRIPT_NAME")


  sSQL = "SELECT * FROM vaEmployee WHERE sEmpno ='" & sEmpno & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sEmpno = Rs("sEmpno")
		sDivcode = Rs("sDivcode")
		sDivision=RS("sDivision")
		sDepcode = Rs("sDepcode")
		sDepartment=RS("sDepartment")
		sSeccode = Rs("sSeccode")
		sSection=Rs("sSection")
		sPoscode = Rs("sPoscode")
		sPosition=RS("sPosition")
		sName = Rs("sName")
		if Instr(sname," ")<>"0" then
			Response.write Mid(sName,1,1) & Lcase(Mid(sname,2,Instr(sname," ")-1)) & Mid(sname,Instr(sname," ")+1,2) & Lcase(Mid(sname,Instr(sname," ")+3,Len(sname)))
		else
			response.write Mid(sname,1,1) & Lcase(Mid(sname,2,Len(sname)))
		End if
	End If
End Sub

Function GetNumber(number)
	Dim Count,i,counter,num(),number1
	Count=Len(number)
	counter=count
	reDim num(count+1)
	For i=1 to Count
		num(i)=mid(number,i,1)
	next 

	IF num(count-1)="." then
		num(count+1)="0"
		count=count+1
		counter=count
	End if
	For i=1 to count
		if i mod 3 =0 and i<> count and num(counter) <> "." then
			number1="," & num(counter) & number1
		else
			number1=num(counter) & number1
		End if

		counter=counter-1
	Next
		response.write number1
End Function

Sub ShowDivision(Divcode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pDivision Where sDivcode='" & Divcode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Response.write Rs2("sDiv")
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub ShowDepartment(DepCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pDepartment Where sDepcode='" & Depcode & "'" &_
		" Order By sDepartment"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
			Do While NOT Rs2.EOF
				Response.write Rs2("sDept")
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub ShowSection(SecCode)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pSection Where sSeccode='" & Seccode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
		IF NOT Rs2.EOF  Then
				Response.write Rs2("sSec")
		End IF
End Sub


Function Show_Budget(prodC, prDate, sExptype)
	Dim sSQL2,RS2
	Dim BgYPR, EdYPr, EPYear
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")	
'	response.write month(prDate)
	if month(prDate) <= 3 then
		BgYPR = year(prDate) - 1
		EdYPr  = year(prDate)
		EPYear = year(prDate) - 1		
	else
		BgYPR = year(prDate)
		EdYPr  = year(prDate) + 1
		EPYear = year(prDate)
	end if

	sSQL2= "select sum(namount) as namount, pcatalog.scatid, nbudget" _
	& " from pproduct, ppritem, pcatalog, epbudget, ppr" _
	& " where pproduct.sprodcode = ppritem.sprodcode" _
	& " and pproduct.scatid = pcatalog.scatid" _
	& " and pcatalog.scatid = epbudget.scatid" _
	& " and ppr.sprno = ppritem.sprno" _
	& " and ppritem.cstatus <> '0' and ppr.sExptype = pproduct.sExptype " _
	& " and ppr.dprdate between to_date('01/04/" & BgYPR & "','dd/mm/yyyy') " _
	& " and to_date('31/03/" & EdYPr &"','dd/mm/yyyy') " _
	& " and epbudget.scatid = (select scatid from pproduct where sprodcode = '" & prodC & "')" _
	& " group by pcatalog.scatid, nbudget"
	
	sSQL2= "select sum(namount) as namount, nbudget" _
	& " from ppritem, pcatalog, epbudget, ppr" _
	& " , (select scatid, sExptype from  pproduct where sprodcode =  '" & prodC & "') prod" _
	& " where prod.scatid = pcatalog.scatid" _
	& " and ppr.sExptype = prod.sExptype" _
	& " and epbudget.scatid = pcatalog.scatid" _
	& " and ppr.sprno = ppritem.sprno" _
	& " and ppritem.cstatus <> '0' " _
	& " and ppr.dprdate between to_date('01/04/" & BgYPR & "','dd/mm/yyyy') " _
	& " and to_date('31/03/" & EdYPr &"','dd/mm/yyyy') " _
	& " group by nbudget" 
	
	sSQL2= "select sum(namount) as namount, nbudget" _
	& " from ppritem, pcatalog, epbudget, ppr, pproduct" _
	& " where pproduct.scatid = pcatalog.scatid" _
	& " and ppr.sExptype = pproduct.sExptype" _
	& " and epbudget.scatid = pcatalog.scatid" _
	& " and ppr.sprno = ppritem.sprno" _
	& " and ppritem.cstatus <> '0' " _
	& " and ppr.dprdate between to_date('01/04/" & BgYPR & "','dd/mm/yyyy') " _
	& " and to_date('31/03/" & EdYPr &"','dd/mm/yyyy') and epbudget.cyear = '" & year(date) & "'" _
	& " and epbudget.scatid = (select pproduct.scatid from  pproduct, epbudget where sprodcode =  '" & prodC & "' and epbudget.scatid = pproduct.scatid  and epbudget.cyear = '" & year(date) & "' and pproduct.sexptype='" & sExptype & "') " _
	& " and pproduct.sprodcode = ppritem.sprodcode " _
	& " group by nbudget" 

	sSQL2= "select sum(namount) as namount, nbudget" _
	& " from ppritem, pcatalog, epbudget, ppr, pproduct" _
	& " where pproduct.scatid = pcatalog.scatid" _
	& " and ppr.sExptype = pproduct.sExptype" _
	& " and epbudget.scatid = pcatalog.scatid" _
	& " and ppr.sprno = ppritem.sprno" _
	& " and ppritem.cstatus <> '0' " _
	& " and ppr.dprdate between to_date('01/04/" & BgYPR & "','dd/mm/yyyy') " _
	& " and to_date('31/03/" & EdYPr &"','dd/mm/yyyy') and epbudget.cyear = '" & EPYear & "'" _
	& " and epbudget.scatid = (select pproduct.scatid from  pproduct, epbudget where sprodcode =  '" & prodC & "' and epbudget.scatid = pproduct.scatid  and epbudget.cyear = '" & EPYear & "' and pproduct.sexptype='" & sExptype & "') " _
	& " and pproduct.sprodcode = ppritem.sprodcode " _
	& " group by nbudget" 
	
'	response.write sSQL2
'response.end
'	& " and ppritem.sprodcode = '" & prodC & "'" _

	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
'	response.write "[" & Rs2.EOF & "]"
	IF NOT Rs2.EOF  Then
		Show_Budget = RS2("nbudget") & "," & RS2("namount")
'		Response.write RS2("nbudget")
'		if cCur(RS2("nbudget")) < cCur(RS2("namount")) then
'			response.write "<font color='#00FFFF'>"
'		else
'			response.write "<font color='#000000'>"		
'		end if
'		Response.write "-" & RS2("namount")
'			response.write "</font>"	
	Else
		Show_Budget = "0,0"
	End if	
End Function

Sub List_Pcatalog(scatid)
Dim sSQL2,RS2,cSelect
	sSQL2= "Select * from pcatalog Order By sCatID"
	'response.write sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("scatid") = scatid Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("scatid")& "'" & cSelect &">"&Rs2("scatid")&"-" & Rs2("scatalog") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

Sub List_Pcattype(scattype)
Dim sSQL2,RS2,cSelect
	sSQL2= "select * from pcattype Order By sCatType"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("scattype") = scattype Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("scattype")& "'" & cSelect &">"&Rs2("scattype")&"-" & Rs2("scatdesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

'-----------------Edit by golf 4 Dec 2017-------------------------------------
Sub List_Hazardno(hazardno)
Dim sSQL2,RS2,cSelect
		sSQL2= "SELECT * FROM HAZARDTYPEMST Order By Hazardno"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("hazardno") = hazardno Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
		Response.Write "<option value= '" & Rs2("hazardno") & "'" & cSelect &">"& Rs2("hazardtype") &"</option>"& vbCrLf
		Rs2.MoveNext
		Loop
	End IF
End Sub
'--------------------------------------------------------------------------------
%>