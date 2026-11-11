<%
function isSCMPO(po)
	Dim rsOra, sqlOra

	set rsOra = createObject("adodb.recordset")
	sqlOra = "select cSentSCM from ppo where spono = '" & po & "' and cSentSCM <> '0' "
	rsOra.open sqlOra, conn
	
	if rsOra.EOF then
		isSCMPO = false
	else
		isSCMPO = true
	end if
	
	rsOra.close
end function

function showRecItemSendSCMStatus(rec, prodCode, conn)
	Dim sqlOra, rsOra

	set rsOra = createObject("adodb.recordset")

	sqlOra = "select dSentSCM, cSentSCMTime " _
	& "from preceiveitem " _
	& "where " _
	& "sRecNo = '" & rec & "' and sProdCode = '" & prodCode & "' "
	'response.write sqlOra
	rsOra.open sqlOra, conn

	if rsOra.EOF then
		showRecItemSendSCMStatus = ""
	else
		showRecItemSendSCMStatus = "<br>Sent to SCM on<br>" & rsOra("dSentSCM") & "," & rsOra("cSentSCMTime")
	end if

	rsOra.close
	set rsOra = nothing
end function

'##########
'# Function: sendRecDataToSCM(po, inv, recDate, prodCode, qty)
'# Description: This function send receiving data to AS/400 in file rtnlibf/k061kp
'# Argument: 
'# 1.po = P/O number
'# 2.inv = invoice number
'# 3.recDate = receiving date
'# 4.prodCode = product code
'# 5.qty = quantity
'# 6.conn = Oracle's connection
'##########
function sendRecDataToSCM(po, rec, inv, recDate, prodCode, qty, nItem, conn)
	Dim rs400, sql400
	Dim sqlOra, rsOra
	Dim sep, rep_sep
	
	'Set seperator
	sep = "|"
	rep_sep = "\"

	set rs400 = createObject("adodb.recordset")
	set rsOra = createObject("adodb.recordset")
	
	sqlOra = "select cSentSCM from ppo where spono = '" & po & "' and cSentSCM <> '0' "
	rsOra.open sqlOra, conn
	
	if rsOra.EOF then
		exit function
	end if
	
	rsOra.close

	'Insert receiving data in rtnlibf.k061kp
	sql400 = "insert into " & rec_export_table _
	& " values ('L" & sep & "1" & sep & po & sep & oraFormatStr(encode(inv, sep, rep_sep)) & sep & toDB2Date(recDate) _
	& sep & oraFormatStr(encode(prodCode, sep, rep_sep)) & sep & qty & sep & "')"
	
	sql400 = "insert into " & rec_export_table _
	& " values ('L" & sep & "1" & sep & po & sep & oraFormatStr(encode(inv, sep, rep_sep)) & sep & toDB2Date(recDate) _
	& sep & oraFormatStr(prodCode) & sep & qty & sep & sep & nItem & "')"
	'response.write sql400 & "<br>"
	rs400.open sql400, connas4002

	'Update sent to SCM status
	sqlOra = "update pReceiveItem " _
	& "set cSentSCM = '1', dSentSCM = to_date('" & recDate & "', 'DD/MM/YY'), cSentSCMTime = '" & hour(time) & ":" & minute(time) & "' " _
	& "where " _
	& "sRecNo = '" & rec & "' and sProdCode = '" & prodCode & "'"
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, conn
end function

function getDefaultShipTo(req, conn)
	Dim sSQL2,RS2,cSelect
	Dim org, sname

		sSQL2="Select sName, aemployee.sSecCode, sSec, aemployee.sDepCode, sDept, aemployee.sDivCode, sDiv " _
		& "from aemployee, psection, pdepartment, pdivision " _
		& "where sEmpNo = '" & req & "' " _
		& "and pSection.sSecCode = aemployee.sSecCode " _
		& "and pDepartment.sDepCode = aemployee.sDepCode " _
		& "and pDivision.sDivCode = aemployee.sDivCode "
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	IF NOT Rs2.EOF  Then
		if rs2("sSecCode") <> "00" then
			org = rs2("sSec")
		else
			if rs2("sDepCode") <> "00" then
				org = rs2("sDept")
			else ' if rs2("sDivCode") <> "00" then
				org = rs2("sDiv")
			end if
		end if

		if instr(rs2("sName"), " ") <> 0 then
'			sname = left(rs2("sName"), instr(rs2("sName"), " ")-1)
			sname = left(rs2("sName"), 70)
		else
			sname = rs2("sName")
		end if

'		getDefaultShipTo = "AMEC (" & left(org, 3) & "," & left(sname, 19) & ")"
		getDefaultShipTo = "AMEC (" & left(org, 3) & "," & left(sname, 70) & ")"
	else
		getDefaultShipTo = "AMEC"
	End IF	

	rs2.close
end function
%>