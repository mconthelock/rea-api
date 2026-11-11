<% 
' ====================================================
'	AbleOn.Com Intranet Solution
'	Common Function

'	Author : Phornthep  Khongsathian
'	Create Date : 13/11/2542
'	Modify By : Phornthep  Khongsathian
'	Modify Date : 24/11/2542
' ====================================================

' --------------------------------------
'Function Convert Date To Every Database
' Oracle 7.3.3
Function To_Date(dDate)
	Dim sDate,sDay,sMonth,sYear

	IF Not IsDate(dDate) Then
		sDate = Date
	Else 
		sDate = dDate
	End IF

	' #### Convert Date 
'response.write "sDate=[" & sDate & "]"
'response.write "sDate=[" & dDate & "]"
'response.end
	' Oracle 7.3.3
'response.write "day=[" & day(sDate) & "]"
'response.write "month=[" & month(sDate) & "]"
'response.write "year=[" & year(sDate) & "]"

 ''   sDay = Day(sDate)
''	sMonth = Month(sDate)
''	sYear = Year(cdate(sDate)) 
'' Modify by Kallayanee on 29/08/2008
	sDay = split(dDate,"/")
''	sYear = Year(cdate(sDate)) +543
''    sDate = sMonth & "/" & sDay& "/" & sYear
	sDate = sDay(1) & "/" & sDay(0) & "/" & sDay (2)
	'response.write "sDate=[" & sDate & "]"	
	To_Date = " to_date('" &sDate &"','MM/DD/YYYY')"
	
	' #### Access 97
	'To_Date = " #" & sDate & "# "
'response.write "day=[" & sDate & "]"
'response.write "day=[" & day(sDate) & "]"
'response.write "month=[" & month(sDate) & "]"
'response.write "year=[" & year(sDate) & "]"
End Function

'===== Convert Date to thai วัน เดือน พ.ศ.========
'Function Convert Date To Show Thai format
Function Show_ThaiDate(dDate)
	Dim sDate,sDay,sMonth,sYear
	
	IF Not IsDate(dDate) Then
		sDate = Date
	Else 
		sDate = dDate
	End IF
	Show_ThaiDate = Day(sDate) &"/"& Month(sDate) &"/"& (Year(sDate)+543)

End Function

'===== Convert Date to thai วัน เดือน พ.ศ.========
'Function Convert Date To Show Thai format
Function Show_ThaiDay(dDate)
	Dim sDate,sDay,sMonth,sYear
	IF Not IsDate(dDate) Then
		sDate = Date
	Else 
		sDate = dDate
	End IF
	Show_ThaiDay = Day(sDate) &" "& MonthName(Month(sDate)) &" "& (Year(sDate)+543)
End Function

'===== Get PR No,PO NO From Sequence ========
Function CreatePRNO
	Dim Rsx,sSQLx,nPOx,ni
	set Rsx= Server.CreateObject("ADODB.RecordSet")
	sSQLx = "SELECT NPRNO.NEXTVAL NPO FROM DUAL"
	Rsx.Open sSQLx, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rsx.EOF  Then
		nPOx = Rsx("NPO")
	End IF
	Rsx.Close
	Set Rsx = Nothing
   If Len(nPOx) > 0 Then
		For ni = 1 to (10 - Len(nPOx))
			nPOx = "0"&nPOx
		next
   end if
   CreatePRNO = "PR"&nPOx
End Function

Function CreatePONO
	Dim Rsx,sSQLx,nPOx,ni
	set Rsx= Server.CreateObject("ADODB.RecordSet")
	sSQLx = "SELECT NPONO.NEXTVAL NPO FROM DUAL"
	Rsx.Open sSQLx, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rsx.EOF  Then
		nPOx = Rsx("NPO")
	End IF
	Rsx.Close
	Set Rsx = Nothing
   If Len(nPOx) > 0 Then
		For ni = 1 to (8 - Len(nPOx))
			nPOx = "0"&nPOx
		next
   end if
   CreatePONO = "AMEC"&nPOx
End Function

Function CreateRecNo
	Dim Rsx,sSQLx,nRecx,ni
	set Rsx= Server.CreateObject("ADODB.RecordSet")
	sSQLx = "SELECT NRECNO.NEXTVAL NREC FROM DUAL"
	Rsx.Open sSQLx, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rsx.EOF  Then
		nRECx = Rsx("NREC")
	End IF
	Rsx.Close
	Set Rsx = Nothing
   If Len(nRECx) > 0 Then
		For ni = 1 to (10 - Len(nRECx))
			nRECx = "0"&nRECx
		next
   end if
   CreateRECNO = "RC"&nRECx
End Function

Function CreatePLANNo
	Dim Rsx,sSQLx,nRecx,ni
	set Rsx= Server.CreateObject("ADODB.RecordSet")
	sSQLx = "SELECT PID.NEXTVAL NPLAN FROM DUAL"
	'response.write sSQLx
	'response.end
	Rsx.Open sSQLx, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rsx.EOF  Then
		nRECx = Rsx("NPLAN")
	End IF
	Rsx.Close
	Set Rsx = Nothing
   If Len(nRECx) > 0 Then
		For ni = 1 to (10 - Len(nRECx))
			nRECx = "0"&nRECx
		next
   end if
   CreatePLANNo = nRECx
End Function

Function CreateTAXNO
	Dim Rsx,sSQLx,nPOx,ni
	set Rsx= Server.CreateObject("ADODB.RecordSet")
	sSQLx = "SELECT NTAXNO.NEXTVAL NPO FROM DUAL"
	Rsx.Open sSQLx, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rsx.EOF  Then
		nPOx = Rsx("NPO")
	End IF
	Rsx.Close
	Set Rsx = Nothing
   If Len(nPOx) > 0 Then
		For ni = 1 to (10 - Len(nPOx))
			nPOx = "0"&nPOx
		next
   end if
   CreateTAXNO = "TX"&nPOx
End Function

Function CreateINVNO
	Dim Rsx,sSQLx,nPOx,ni
	set Rsx= Server.CreateObject("ADODB.RecordSet")
	sSQLx = "SELECT NINVNO.NEXTVAL NPO FROM DUAL"
	Rsx.Open sSQLx, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rsx.EOF  Then
		nPOx = Rsx("NPO")
	End IF
	Rsx.Close
	Set Rsx = Nothing
   If Len(nPOx) > 0 Then
		For ni = 1 to (10 - Len(nPOx))
			nPOx = "0"&nPOx
		next
   end if
   CreateINVNO = "IV"&nPOx
End Function

'================ Check excite Data ==================
' --------------------------------------
' Check excite ,return suserid
Function check_user(sEmpno)
	Dim Rs2,sSQL2
	 ' Get Project Function , Status
	set Rs2= Server.CreateObject("ADODB.RecordSet")
    sSQL2 = "SELECT sEmpno FROM  aEmployee WHERE sEmpno='" & sEmpno &"'"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		check_user = Rs2("sEmpno")
	Else
		check_user = ""
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Function

'================ Get Data Description ==================
' --------------------------------------
'Retrieve data from aCorp,return Core Name description(scorpname)
Function Get_Program(sID)
	Dim Rs2,sSQL2
	 ' Get Project Function , Status
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM aprogram Where sprogcode='" &sID &"'"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Get_Program = Rs2("sprogDesc")
	Else
		Get_Program = ""
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Function

' --------------------------------------
'Retrieve data from aOpport,return Oppdesc description(soppdesc)
Function Get_CatTdesc(sID)
	Dim Rs2,sSQL2
	 ' Get Project Function , Status
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM pCATTYPE Where sCATTYPE='" &sID &"'"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Get_CatTdesc = Rs2("sCATdesc")
	Else
		Get_CatTdesc = ""
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Function

'================= Create Combobox Data ==================
'----------------------------------
'Create Combobox data from corptype,return corp type  description(stypedesc)
Sub create_corptypelist(scoprtype)
	Dim Rs2,sSQL2,cSelect
	 ' Get Program
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM acorptype"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("scorptype") = scorptype Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("scorptype")& "'" & cSelect &">"& Rs2("stypedesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Sub

'--------------------------------------
'Create hour list
Sub Create_hourList(nHour)
	 'Insert Month
		Dim sSelect,ihour,shour,svhour,hhour
		For ihour = 1 to 24 
    		 For hhour = 1 to 2 
				If iHour < 10 Then
					if hhour = 1 then
						sHour = "0"&iHour&":00"
						svHour = "0"&iHour&".00"
					else
						sHour = "0"&iHour&":30"
						svHour = "0"&iHour&".30"
					end if
				Else
					if hhour = 1 then					
						sHour = iHour&":00"
						svHour = iHour&".00"
					else
        				if ihour = 24 then
							sHour = "00:30"
							svHour = "00.30"
    					else
							sHour = iHour&":30"
							svHour = iHour&".30"
						end if
					end if
				End IF
				If nhour = svhour Then
					sSelect = "selected"
				Else
					sSelect = ""
				End IF
				Response.Write "<option value= "&svhour& " " & sSelect &">"& sHour &"</option>"
			next
		Next
End Sub

'--------------------------------------
'Create day list
Sub Create_DayList(nday)
	 'Insert Month
		Dim sSelect,iday
		 Response.Write "<option value=0></option>"
			For iday = 1 to 31 
			If iday = nday Then
				sSelect = "selected"
			Else
				sSelect = ""
			End IF
			Response.Write "<option value= "&iday& " " & sSelect &">"& (iday) &"</option>"
		Next
End Sub

'--------------------------------------
'Create Month list
Sub Create_MonthList(nMonth)
	 'Insert Month
		Dim sSelect,iMonth
		 Response.Write "<option value=0></option>"
			For iMonth = 1 to 12 
			If iMonth = nMonth Then
				sSelect = "selected"
			Else
				sSelect = ""
			End IF
			Response.Write "<option value= "&iMonth& " " & sSelect &">"& getMonthName(iMonth) &"</option>"
		Next
End Sub

'----------------------------------
'Create Combobox data from aopport,return opport description(soppdesc)
Sub create_CATTlist(sccode)
	Dim oSelect,sSQLo,Rso
	set Rso= Server.CreateObject("ADODB.RecordSet")
	sSQLo = "SELECT * FROM pCATTYpe where cstatus =1 order by sexptype, scattype"
	Rso.Open sSQLo, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rso.EOF  Then
		Do While NOT Rso.EOF
			If Rso("SCATTYPE") = sccode Then
				oSelect = "selected"
			Else
				oSelect = ""
			End IF
			Response.Write "<option value= '"&Rso("SCATTYPE")& "'" & oSelect &">"& Rso("SCATDESC") &"</option>"& vbCrLf
			Rso.MoveNext
		Loop
	End IF
	Rso.Close
	Set Rso = Nothing
End Sub

'----------------------------------
'Create Combobox data from owner budget
Sub create_ownerbudget(vorgno)
	Dim oSelect,sSQLo,Rso
	set Rso= Server.CreateObject("ADODB.RecordSet")
	sSQLo = "SELECT * FROM OWNERBUDGET ORDER BY VORGNAME"
	Rso.Open sSQLo, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rso.EOF  Then
		Do While NOT Rso.EOF
			If Rso("VORGNO") = vorgno Then
				oSelect = "selected"
			Else
				oSelect = ""
			End IF
			Response.Write "<option value= '"&Rso("VORGNO")& "'" & oSelect &">"& Rso("VORGNAME") &"</option>"& vbCrLf
			Rso.MoveNext
		Loop
	End IF
	Rso.Close
	Set Rso = Nothing
End Sub

'----------------------------------
'Create Combobox data from aprogram,return sprogdesc description(sprogdesc)
Sub create_proglist(sprogcode)
	Dim Rs2,sSQL2,cSelect
	 ' Get Program
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM aprogram"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("sprogcode") = sprogcode Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("sprogcode")& "'" & cSelect &">"& Rs2("sprogdesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Sub

'----------------------------------
'Create Combobox data from pPosition
Sub create_pPosition(sPosition)
	Dim Rs2,sSQL2,cSelect
	 ' Get Program
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM pPosition order by sPoscode"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("sPoscode") = sPosition Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("sPoscode")& "'" & cSelect &">"& Rs2("sPosition") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Sub

'----------------------------------
'Create Combobox data from pSection
Sub create_pSection(sSection)
	Dim Rs2,sSQL2,cSelect
	 ' Get Program
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM pSection order by sSeccode"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("sSeccode") = sSection Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("sSeccode")& "'" & cSelect &">"& Rs2("sSection") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Sub

'----------------------------------
'Create Combobox data from pDepartment
Sub create_pDepartment(sDepartment)
	Dim Rs2,sSQL2,cSelect
	 ' Get Program
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM pDepartment order by sDepcode "
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("sDepcode") = sDepartment Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("sDepcode")& "'" & cSelect &">"& Rs2("sDepartment") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Sub

'----------------------------------
'Create Combobox data from pDivision
Sub create_pDivision(sDivision)
	Dim Rs2,sSQL2,cSelect
	 ' Get Program
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM pDivision order by sDivcode"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("sDivcode") = sDivision Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("sDivcode")& "'" & cSelect &">"& Rs2("sDivision") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Sub

'----------------------------------
'Create Combobox data from pStatus
Sub create_pStatus(sStatus)
	Dim Rs2,sSQL2,cSelect
	 ' Get Program
	set Rs2= Server.CreateObject("ADODB.RecordSet")
	sSQL2 = "SELECT * FROM pStatusEmp order by sStatuscode"
	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		Do While NOT Rs2.EOF
			If Rs2("sStatuscode") = sStatus Then
				cSelect = "selected"
			Else
				cSelect = ""
			End IF
			Response.Write "<option value= '"&Rs2("sStatuscode")& "'" & cSelect &">"& Rs2("sStatusDesc") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	Rs2.Close
	Set Rs2 = Nothing
End Sub
%>