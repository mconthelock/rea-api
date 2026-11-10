<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include virtual="/sys/inc/numeric.inc"-->

<%
Dim rsOra, sqlOra

set rsOra = createObject("adodb.recordset")

connOra.beginTrans

sqlOra = "insert into sequenceOrg " _
& "(empNo, vOrgNo, sPosCode, cCo, headNo, vOrgNo1, sPosCode1, cCo1) " _
& "values " _
& "('" & oraFormatStr(request("txtEmpNo")) & "', '" & request("selOrg") & "', '" & request("selPos") & "', '" & toInt(request("chkCo")) & "', '" & oraFormatStr(request("txtMgrNo")) & "', '" & request("selOrg1") & "', '" & request("selPos1") & "', '" & toInt(request("chkCo1")) & "')"
'response.write sqlOra & "<BR>"

on error resume next
rsOra.open sqlOra, connOra
verifyData

if connOra.errors.count > 0 then
	connOra.rollbackTrans
	
	'Validate data
	validateData
	
	'showOraError connOra.errors(0).description & "&&but_back"
	response.end
else
	connOra.commitTrans
end if

on error goto 0

set rsOra = nothing
response.redirect "../orgFrm.asp"

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