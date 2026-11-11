<%
'Add by Dermpan on 30-12-03 due to the SCM system preparation for implementation.
dim export_table, export_field, rec_export_table, rec_export_field
dim exp_ora_table, exp_ora_field

export_table = "rtnlibf.k050kp" 
export_field = "k050kp"
exp_ora_table = "amec.expscm"
exp_ora_field = "scmdata"

'export_table = "rtnlibf.k090kp" 
'export_field = "DTL"

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

	For i=2000 to year(date)+2
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
'	response.write sqlOra
	rsOra.open sqlOra, conn
	
	alreadySentSCM = not rsOra.EOF
	
	rsOra.close
end function

function isInSCMTray(poNo)
	dim rs400, sql400
	
	set rs400 = createObject("adodb.recordset")

	sql400 = "select * from " & export_table & " where " & export_field & " like 'H%" & poNo & "%' or " & export_field & " like 'L%" & poNo & "%'"
	'response.write sql400
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
	
	sqlOra = "delete from " & exp_ora_table & " where " & exp_ora_field  & " like 'H%" & poNo & "%' or " & exp_ora_field & " like 'L%" & poNo & "%'"
	rsOra.open sqlOra, conn
	
	'Update flag
	sqlOra = "update ppo set cSentScm = '0' where sPoNo = '" & poNo & "'"
	rsOra.open sqlOra, conn
end function

function sendToSCM(poNo)
	dim rsOra, rsOra2, rsOra3, sqlOra, rs400, sql400
	dim sep, rep_sep, i, sep1, stype
	dim dpodate, ddelivdate
	
	set rsOra = createObject("adodb.recordset")
	set rsOra2 = createObject("adodb.recordset")
	set rsOra3 = createObject("adodb.recordset")	
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
	sep1 = "|"
	rep_sep = "\"

    If rsOra.EOF Then
	rsOra.close
    Else
		'Delete currenty P/O data
		sql400 = "delete from " & export_table & " " _
		& "where " _
		& export_field & " like 'H" & sep & rsOra("sPoNo") & "%' or " & export_field & " like 'L" & sep & rsOra("sPoNo") & "%'"
		rs400.open sql400, connas4002
		
		sqlOra = "delete from " & exp_ora_table & " " _
		& "where " _
		& exp_ora_field & " like 'H" & sep1 & rsOra("sPoNo") & "%' or " & exp_ora_field & " like 'L" & sep1 & rsOra("sPoNo") & "%'"
		rsOra2.Open sqlOra, conn
		
        Do While Not rsOra.EOF
            i = 1
            '########## READ PO ITEM DATA ##########
            sqlOra = "select ppoitem.*, to_number(ppoitem.nQuality) as nQuality2, to_number(ppoitem.nPrice) as nPrice2, sEProdName, pproduct.sExpType , punit.sunitdesc " _
            & "from ppo, ppr, ppoitem, pproduct " _
            & "where ((ppo.sPoNo = '" & rsOra("sPoNo") & "' and ppoitem.spono = ppo.spono) and ppr.sprno = srefno) " _
            & "and pproduct.sexptype = ppr.sexptype and pproduct.sprodcode = ppoitem.sprodcode "
			
			sqlOra = "select ppoitem.*, to_char(ppoitem.nQuality) as nQuality2, to_char(ppoitem.nPrice) as nPrice2, sEProdName, pproduct.sExpType , punit.sunitdesc " _
			& ", replace(ppoitem.sProdName, '|', '-') as productname " _
            & "from ppo, ppr, ppoitem, pproduct , punit " _
            & "where ((ppo.sPoNo = '" & rsOra("sPoNo") & "' and ppoitem.spono = ppo.spono) and ppr.sprno = srefno) " _
            & "and pproduct.sexptype = ppr.sexptype and pproduct.sprodcode = ppoitem.sprodcode and ppoitem.sunitcode = punit.sunitcode "
			
			'response.write "sqlOra = "&sqlOra
			'response.end
            rsOra2.Open sqlOra, conn
            'If conn.Errors.Count > 0 Then GoTo errorHandle
			
            If rsOra2.EOF Then
            Else
                'Write P/O data
				dpodate=day(FormatDateTime(rsOra("dPoDate"), 2)) & "/" & month(FormatDateTime(rsOra("dPoDate"), 2)) & "/" & year(FormatDateTime(rsOra("dPoDate"), 2))
				ddelivdate=day(FormatDateTime(rsOra("dDelivery"), 2)) & "/" & month(FormatDateTime(rsOra("dDelivery"), 2)) & "/" & year(FormatDateTime(rsOra("dDelivery"), 2))
				if isSCMVender(rsOra("sVendCode"), conn) then
					stype = "1"
				else
					stype = "2"
				end if
				
                sql400 = "insert into " & export_table & " values ('H" & sep & rsOra("sPoNo") & sep & stype & sep & "2" & sep & toDB2Date(dpodate) & sep _
                & oraFormatStr(encode(rsOra("sVendCode"), sep, rep_sep)) & sep & oraFormatStr(encode(rsOra("sRefNo"), sep, rep_sep)) & sep _
                & oraFormatStr(encode(rsOra("sShipTo"), sep, rep_sep)) & sep & oraFormatStr(encode(rsOra("sShipBy"), sep, rep_sep)) & sep & oraFormatStr(encode(rsOra("sPayTerm"), sep, rep_sep)) & sep _
                & oraFormatStr(encode(rsOra("sTradeTerm"), sep, rep_sep)) & sep & oraFormatStr(Mid(rsOra("sByer"), InStr(rsOra("sByer"), "(") + 1, 5)) & sep & sep _
                & sep & toDB2Date(ddelivdate) & sep & oraFormatStr(encode(getCurCode(rsOra("sPoNo")), sep, rep_sep)) & "')"
				'	response.write sql400 & "<br>"

				''''rs400.open sql400, connas4002
				
                sqlOra = "insert into " & exp_ora_table & " (scmdata,id, spono) values ('H" & sep1 & rsOra("sPoNo") & sep1 & "3" & sep1 & stype & sep1 & toDB2Date(dpodate) & sep1 _
                & oraFormatStr(encode1(rsOra("sVendCode"), sep1, rep_sep)) & sep1 & oraFormatStr(encode1(rsOra("sRefNo"), sep1, rep_sep)) & sep1 _
                & oraFormatStr(encode1(rsOra("sShipTo"), sep1, rep_sep)) & sep1 & oraFormatStr(encode1(rsOra("sShipBy"), sep1, rep_sep)) & sep1 & oraFormatStr(encode1(rsOra("sPayTerm"), sep1, rep_sep)) & sep1 _
                & oraFormatStr(encode1(rsOra("sTradeTerm"), sep1, rep_sep)) & sep1 & oraFormatStr(Mid(rsOra("sByer"), InStr(rsOra("sByer"), "(") + 1, 5)) & sep1 & sep1 _
                & sep1 & toDB2Date(ddelivdate) & sep1 & oraFormatStr(encode1(getCurCode(rsOra("sPoNo")), sep1, rep_sep)) &  sep1 & "N','0','" & rsOra("sPoNo") & "')"
				rsOra3.Open sqlOra, conn
				
                Do While Not rsOra2.EOF
                    'Write P/O item data
                    sql400 = "insert into " & export_table & " values ('L" & sep & encode(rsOra2("sPoNo"), sep, rep_sep) & sep & i & sep & sep & sep & sep _
                    & oraFormatStr(encode(rsOra2("sProdName"), sep, rep_sep)) & sep & oraFormatStr(rsOra2("sProdCode")) & sep _
                    & sep & sep & sep & sep _
                    & sep & rsOra2("nQuality2") & sep & oraFormatStr(encode(rsOra2("sUnitCode"), sep, rep_sep)) _
                    & sep & rsOra2("nPrice2") & sep & sep & "')"
					''''rs400.open sql400, connas4002
					
                    sqlOra = "insert into " & exp_ora_table & " (scmdata,id, spono) values ('L" & sep1 & encode1(rsOra2("sPoNo"), sep1, rep_sep) & sep1 & i & sep1 & sep1 & sep1 & sep1 _
                    & oraFormatStr(encode1(rsOra2("productname"), sep1, rep_sep)) & sep1 & oraFormatStr(rsOra2("sProdCode")) & sep1 _
                    & sep1 & sep1 & sep1 & sep1 _
                    & sep1 & rsOra2("nQuality2") & sep1 & oraFormatStr(encode1(rsOra2("sunitdesc"), sep1, rep_sep)) _
                    & sep1 & rsOra2("nPrice2") & sep1 & sep1 & sep1 & toDB2Date(ddelivdate) & "','" & i & "','" & rsOra2("sPoNo") & "')"
					rsOra3.Open sqlOra, conn
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

function decode1(txt, sep1, rep_sep)
	decode1 = replace(replace(txt & "", rep_sep & rep_sep, rep_sep), rep_sep & sep1, sep1)
end function

Function encode(txt, sep, rep_sep)
    encode = trim(replace(replace(txt & "", rep_sep, rep_sep & rep_sep), sep, rep_sep & sep))
End Function
Function encode1(txt, sep1, rep_sep)
    encode1 = trim(replace(replace(txt & "", rep_sep, rep_sep & rep_sep), sep1, rep_sep & sep1))
End Function


Function toDB2Date(vdate)
 ''   toDB2Date = String(4 - Len(Year(vdate)), "0") & Year(vdate) & String(2 - Len(Month(vdate)), "0") & Month(vdate) & String(2 - Len(Day(vdate)), "0") & Day(vdate)
    toDB2Date = String(4 - Len(Year(vdate)), "0") & right(vdate,4) & String(2 - Len(Month(vdate)), "0") & Month(vdate) & String(2 - Len(Day(vdate)), "0") & Day(vdate)
''   toDB2Date = String(4 - Len(Year(vdate)), "0") & (right(vDate,4)-543) & String(2 - Len(Month(vdate)), "0") & Month(vdate) & String(2 - Len(Day(vdate)), "0") & Day(vdate)
	
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