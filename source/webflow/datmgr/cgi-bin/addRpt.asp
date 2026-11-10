<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/date.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no, runNo, start

set rsOra = createObject("adodb.recordset")

'Validate data
validateReportData

connOra.beginTrans

'Get new report number
sqlOra = "select max(nRepNo) from rptMst where cYear = '" & right(year(date), 2) & "'"
rsOra.open sqlOra, connOra

if isNull(rsOra("max(nRepNo)")) then
	no = 1
else
	no = cInt(rsOra("max(nRepNo)")) + 1
end if
rsOra.close

'Insert new step
sqlOra = "insert into rptMst (nRepNo, cYear, vName, vDesc, vUrl, vOrgNo, cStatus) " _
& "values " _
& "(" & no & ", '" & right(year(date), 2) & "', '" & oraFormatStr(left(trim(request.queryString("txtName")), 120)) & "', '" & oraFormatStr(left(trim(request.queryString("txtDesc")), 256)) & "', '" & oraFormatStr(left(trim(request("txtUrl")), 1024)) & "', '" & request.queryString("selOrg") & "', '" & request("optShow") & "')"
'response.write sqlOra & "<br>"
'response.end
on error resume next
rsOra.open sqlOra, connOra
showOraError connOra

sqlOra = "insert into rptUsrAuth (nFrmNo, vOrgNo, cYear, vEmpNo, cAuthNo) select '" & no & "', '" & request("selOrg") & "', '" & right(year(date), 2) & "', sEmpNo, '" & AUTH_RPT_READ & "' from amec.aEmployee where amec.aEmployee.cStatus = '1'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra
showOraError connOra

on error goto 0
connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../rptLst.asp"

function validateReportData()
	if trim(request.queryString("selOrg")) = "" then
		response.write "Organize not found!"
		response.end
	end if
	
	if trim(request.queryString("txtName")) = "" then
		response.write "Report's name not found!"
		response.end
	end if
	
	if trim(request.queryString("txtUrl")) = "" then
		response.write "Report's URL not found!"
		response.end
	end if

	'Check for duplicated
	sqlOra = "select nRepNo from rptMst where vOrgNo = '" & request.queryString("selOrg") & "' and lower(vName) = '" & oraFormatStr(lCase(request.queryString("txtName"))) & "'"
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, connOra
	
	if not rsOra.EOF then
		response.write "Duplicated report!"
		response.end
	end if
	
	rsOra.close
end function
%>