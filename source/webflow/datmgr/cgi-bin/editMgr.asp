<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include virtual="/sys/inc/numeric.inc"-->

<%
Dim rsOra, sqlOra

set rsOra = createObject("adodb.recordset")

connOra.beginTrans

sqlOra = "update sequenceOrg set " _
& "empNo = " & translateNullField(oraFormatStr(request("txtEmpNo"))) & ", vOrgNo = " & translateNullField(request("selOrg")) & ", sPosCode = " & translateNullField(request("selPos")) & " " _
& ", cCo = " & translateNullField(toInt(request("chkCo"))) & ", headNo = " & translateNullField(oraFormatStr(request("txtMgrNo"))) & ", vOrgNo1 = " & translateNullField(request("selOrg1")) & " " _
& ", sPosCode1 = " & translateNullField(request("selPos1")) & ", cCo1 = " & translateNullField(toInt(request("chkCo1"))) & " " _
& "where " _
& "empNo " & translateFieldCompareClause(oraFormatStr(request("empNo"))) & " and vOrgNo " & translateFieldCompareClause(request("orgNo")) & " and sPosCode " & translateFieldCompareClause(request("pos")) & " " _
& "and cCo " & translateFieldCompareClause(request("co")) & " and headNo " & translateFieldCompareClause(oraFormatStr(request("headNo"))) & " and vOrgNo1 " & translateFieldCompareClause(request("orgNo1")) & " " _
& "and sPosCode1 " & translateFieldCompareClause(request("pos1")) & " and cCo1 " & translateFieldCompareClause(request("co1")) & " "
'response.write sqlOra & "<BR>"

on error resume next
rsOra.open sqlOra, connOra
verifyData

if connOra.errors.count <> 0 then
	connOra.rollbackTrans
	
	'Validate data
	validateData
	
	'showOraError connOra.errors(0).description & "&&but_back"
	response.write connOra.errors(0).description
	response.end
else
	connOra.commitTrans
end if

on error goto 0

set rsOra = nothing
response.redirect "../mgrList.asp" & "?selDiv=" & request("selDiv") & "&selDep=" & request("selDep") & "&selSec=" & request("selSec")

function verifyData()
	Dim rsOra, sqlOra
	
	set rsOra = createObject("adodb.recordset")
	
	'Check recursively link
	sqlOra = "select * from sequenceOrg " _
	& "start with empNo = '" & request("txtEmpNo") & "' and sPosCode = '" & request("selPos") & "' " _
	& "connect by prior headNo = empNo and prior sPosCode1 = sPosCode "
	'response.write sqlOra & "<br>"
	on error resume next
	rsOra.open sqlOra, connOra
	
	if connOra.errors.count > 0 then
		'response.write connOra.errors(0).description
		response.write "Found recursivly link (loop), please check relation between employee and manager!"
		connOra.rollbackTrans
		response.end
	end if
	
	on error goto 0
end function

function validateData
	if trim(request("txtEmpNo")) = "" then
		response.write "Employee number not found!"
	end if
	
	if trim(request("txtMgrNo")) = "" then
		response.write "Manager's employee number not found!"
	end if
end function
%>