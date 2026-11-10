<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/date.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, pos, nextStep, start

set rsOra = createObject("adodb.recordset")

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Set root flag
if request.queryString("selPrevStep") = "" then
	start = "1"
else
	start = "0"
end if

'Set node next
if request.queryString("selPrevStep") = "" then
	nextStep = "00"
else
	sqlOra = "select distinct cStepNextNo from orgFlowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("selPrevStep") & "'"
	rsOra.open sqlOra, connOra
	
	if not rsOra.EOF then
		nextStep = rsOra("cStepNextNo")
	else
		nextStep = "00"
	end if
	
	rsOra.close
end if

if request.queryString("selPos") = "" then
	pos = "null"
else
	pos = "'" & request.queryString("selPos") & "'"
end if

'Insert new step
sqlOra = "insert into orgFlowMst (nFrmNo, vOrgNo, cYear, cStepNo, cStepNextNo, vUrl, vPosNo, cStart, cType) " _
& "values (" & request.queryString("no") & ", '" & request.queryString("orgNo") & "', '" & request.queryString("y") & "', '" & request.queryString("selStep") & "', '" & nextStep & "', '" & request.queryString("txtUrl") & "', " & pos & ", '" & start & "', '2')"
'response.write sqlOra & "<br>"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if

'Update previous step
sqlOra = "update orgFlowMst set cStepNextNo = '" & request.queryString("selStep") & "' where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("selPrevStep") & "'"
rsOra.open sqlOra, connOra

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../editFrmFrm.asp?no=" & request.queryString("no") & "&orgNo=" & request.queryString("orgNo") & "&y=" & request.queryString("y")
%>