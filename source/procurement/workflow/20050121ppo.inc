<%
'Add by Dermpan on 30-12-03 due to the SCM system preparation for implementation.
dim export_table, export_field, rec_export_table, rec_export_field

export_table = "rtnlibf.k050kp" 
export_field = "k050kp"
rec_export_table = "rtnlibf.k061kp"
rec_export_field = "k61dta"

function getEmpName(empNo)
	dim rsOra, sqlOra

	set rsOra = createObject("adodb.recordset")
	sqlOra = "select sName from aEmployee where sEmpNo = '" & empNo & "'"
	rsOra.open sqlOra, conn

	if rsOra.EOF then
		getEmpName = ""
	else
		getEmpName = rsOra("sName")
	end if

	rsOra.close
end function

function isSCMVender(vender, conn)
	dim rsOra, sqlOra
	
	set rsOra = createObject("adodb.recordset")
	sqlOra = "select * from scmVender where svendcode = '" & vender & "'"
	rsOra.open sqlOra, conn
	isSCMVender = not rsOra.EOF
	rsOra.close
end function

Sub Create_Yearlist(searchyear)
Dim syear, i
	
	if isNull(searchyear) then
		searchyear = year(date)
	end if
	
	response.write "<option value=""""></option>"

	For i=2000 to year(date)+1
		IF i=cInt(searchyear) then
			Response.Write "<option value= '"& i & "'selected>"& i &"</option>"& vbCrLf
		Else
			Response.Write "<option value= '"& i & "'>"& i &"</option>"& vbCrLf
		End IF
	Next
End Sub

function getImpDate(poNo)
	dim rsOra, sqlOra
	
	set rsOra = createObject("adodb.recordset")
	sqlOra = "select dImpSCM from amec.ppo where sPoNo = '" & poNo & "' and cSentSCM = '2'"
	rsOra.open sqlOra, conn
	
	if rsOra.EOF then
		getImpDate = not rsOra.EOF
	else
		getImpDate = rsOra("dImpSCM")
	end if
	
	rsOra.close
end function

function updatedBySCM(poNo)
	dim rsOra, sqlOra
	
	set rsOra = createObject("adodb.recordset")
	sqlOra = "select cSentSCM from amec.ppo where sPoNo = '" & poNo & "' and cSentSCM = '2'"
	rsOra.open sqlOra, conn
	
	updatedBySCM = not rsOra.EOF
	
	rsOra.close
end function

function alreadySentSCM(poNo)
	dim rsOra, sqlOra
	
	set rsOra = createObject("adodb.recordset")
	sqlOra = "select cSentSCM from amec.ppo where sPoNo = '" & poNo & "' and cSentSCM = '1'"
	rsOra.open sqlOra, conn
	
	alreadySentSCM = not rsOra.EOF
	
	rsOra.close
end function

function isInSCMTray(poNo)
	dim rs400, sql400
	
	set rs400 = createObject("adodb.recordset")

	sql400 = "select * from " & export_table & " where " & export_field & " like 'H%" & poNo & "%' or " & export_field & " like 'L%" & poNo & "%'"
	rs400.open sql400, connas4002
	
	isInSCMTray = not rs400.EOF
	
	rs400.close
end function


function removeFromSCM(poNo)
	dim rsOra, sqlOra, rs400, sql400
	
	set rs400 = createObject("adodb.recordset")
	set rsOra = createObject("adodb.recordset")
	
	'Remove from SCM
	sql400 = "delete from " & export_table & " where " & export_field & " like 'H%" & poNo & "%' or " & export_field & " like 'L%" & poNo & "%'"
	rs400.open sql400, connas4002
	
	'Update flag
	sqlOra = "update ppo set cSentScm = '0' where sPoNo = '" & poNo & "'"
	rsOra.open sqlOra, conn
end function

function sendToSCM(poNo)
	dim rsOra, rsOra2, sqlOra, rs400, sql400
	dim sep, rep_sep, i

	set rsOra = createObject("adodb.recordset")
	set rsOra2 = createObject("adodb.recordset")
	set rs400 = createObject("adodb.recordset")

	if poNo = "" or isNull(poNo) or isEmpty(poNo) then
		exit function
	end if
	
    '########## READ PO DATA ##########
    sqlOra = "select ppo.*, pVender.sEAddress, pVender.sEName, sTelNo " _
    & "from ppo, pVender " _
    & "where ppo.cStatus <> '9' and pVender.sVendCode = ppo.sVendCode and sPoNo = '" & poNo & "'"
    rsOra.Open sqlOra, Conn
   ' If conn.Errors.Count > 0 Then GoTo errorHandle

	'Set seperator
	sep = "%"
	rep_sep = "\"

    If rsOra.EOF Then
	rsOra.close
    Else
		'Delete currenty P/O data
		sql400 = "delete from " & export_table & " " _
		& "where " _
		& export_field & " like 'H" & sep & rsOra("sPoNo") & "%' or " & export_field & " like 'L" & sep & rsOra("sPoNo") & "%'"
		rs400.open sql400, connas4002
		
        Do While Not rsOra.EOF
            i = 1
            '########## READ PO ITEM DATA ##########
            sqlOra = "select ppoitem.*, to_number(ppoitem.nQuality) as nQuality2, to_number(ppoitem.nPrice) as nPrice2, sEProdName, pproduct.sExpType " _
            & "from ppo, ppr, ppoitem, pproduct " _
            & "where ((ppo.sPoNo = '" & rsOra("sPoNo") & "' and ppoitem.spono = ppo.spono) and ppr.sprno = srefno) " _
            & "and pproduct.sexptype = ppr.sexptype and pproduct.sprodcode = ppoitem.sprodcode "
            rsOra2.Open sqlOra, conn
            'If conn.Errors.Count > 0 Then GoTo errorHandle

            If rsOra2.EOF Then
            Else
                'Write P/O data
                sql400 = "insert into " & export_table & " values ('H" & sep & rsOra("sPoNo") & sep & "3" & sep & "2" & sep & toDB2Date(rsOra("dPoDate")) & sep _
                & oraFormatStr(encode(rsOra("sVendCode"), sep, rep_sep)) & sep & oraFormatStr(encode(rsOra("sRefNo"), sep, rep_sep)) & sep _
                & oraFormatStr(encode(rsOra("sShipTo"), sep, rep_sep)) & sep & oraFormatStr(encode(rsOra("sShipBy"), sep, rep_sep)) & sep & oraFormatStr(encode(rsOra("sPayTerm"), sep, rep_sep)) & sep _
                & oraFormatStr(encode(rsOra("sTradeTerm"), sep, rep_sep)) & sep & oraFormatStr(Mid(rsOra("sByer"), InStr(rsOra("sByer"), "(") + 1, 5)) & sep & sep _
                & sep & toDB2Date(rsOra("dDelivery")) & sep & oraFormatStr(encode(getCurCode(rsOra("sPoNo")), sep, rep_sep)) & "')"
	'response.write sql400 & "<br>"
				rs400.open sql400, connas4002
				
                Do While Not rsOra2.EOF
                    'Write P/O item data
                    sql400 = "insert into " & export_table & " values ('L" & sep & encode(rsOra2("sPoNo"), sep, rep_sep) & sep & i & sep & sep & sep & sep _
                    & oraFormatStr(encode(rsOra2("sEProdName"), sep, rep_sep)) & sep & oraFormatStr(rsOra2("sProdCode")) & sep _
                    & sep & sep & sep & sep _
                    & sep & rsOra2("nQuality2") & sep & oraFormatStr(encode(rsOra2("sUnitCode"), sep, rep_sep)) _
                    & sep & rsOra2("nPrice2") & sep & sep & "')"
					rs400.open sql400, connas4002
                    i = i + 1
	'response.write sql400 & "<br>"
                    rsOra2.MoveNext
                Loop
            End If
            
            rsOra2.Close
            rsOra.MoveNext
            '########## END READ PO ITEM DATA ##########
        Loop
        '########## END READ PO DATA ##########

	
	rsOra.Close
	
	sqlOra = "update amec.ppo set cSentScm = '1', " _
	& "dSentSCM = to_date('" & Day(Date) & "/" & Month(Date) & "/" & Year(Date) & "', 'DD/MM/YY') " _
	& ", cSentSCMTime = '" & hour(time) & ":" & minute(time) & "' " _
	& " where sPoNo = '" & poNo & "'"
	rsOra.open sqlOra, conn
    End If
	
	set rsOra = nothing
end function

function decode(txt, sep, rep_sep)
	decode = replace(replace(txt & "", rep_sep & rep_sep, rep_sep), rep_sep & sep, sep)
end function

Function encode(txt, sep, rep_sep)
    encode = trim(replace(replace(txt & "", rep_sep, rep_sep & rep_sep), sep, rep_sep & sep))
End Function

Function toDB2Date(vdate)
    toDB2Date = String(4 - Len(Year(vdate)), "0") & Year(vdate) & String(2 - Len(Month(vdate)), "0") & Month(vdate) & String(2 - Len(Day(vdate)), "0") & Day(vdate)
End Function

Function getUnitCode(poNo, prodCode)
    Dim rsOra, sqlOra
    
	set rsOra = createObject("adodb.recordset")
	
    sqlOra = "select vCode from amec.exccode where vEPCode in (select sUnitCode from ppoitem where sPoNo = '" & poNo & "' and sProdCode = '" & prodCode & "')"
    rsOra.Open sqlOra, conn
    
    If rsOra.EOF Then
        getUnitCode = ""
    Else
        getUnitCode = rsOra("vCode")
    End If
    
    rsOra.Close
End Function

Function getCurCode(poNo)
    Dim rsOra, sqlOra
    
	set rsOra = createObject("adodb.recordset")
	
    sqlOra = "select vCode from amec.exccode where vEPCode in (select sCurCode from ppoitem where sPoNo = '" & poNo & "')"
    rsOra.Open sqlOra, conn
    
    If rsOra.EOF Then
        getCurCode = ""
    Else
        getCurCode = rsOra("vCode")
    End If
    
    rsOra.Close
End Function

function timeIsOK()
	if time() >= "12:10" and time() <= "13:15" then
		timeIsOK = false
	else
		timeIsOK = true
	end if
end function

'########## COPY AND MODIFY FROM ../fcommon.asp ##########
Sub SCM_DayList(nday)
	 'Insert Month
		Dim sSelect,iday
			response.write "<option value=""""></option>"
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
Sub SCM_MonthList(nMonth)
	 'Insert Month
		Dim sSelect,iMonth
			response.write "<option value=""""></option>"
			For iMonth = 1 to 12 
			If iMonth = nMonth Then
				sSelect = "selected"
			Else
				sSelect = ""
			End IF
			Response.Write "<option value= "&iMonth& " " & sSelect &">"& getMonthName(iMonth) &"</option>"
		Next
End Sub
'########## END COPY ##########
%>