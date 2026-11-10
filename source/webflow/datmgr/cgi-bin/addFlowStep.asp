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
Dim apv, apvOrg, applyAll

set rsOra = createObject("adodb.recordset")

'Check existing step
sqlOra = "select nFrmNo from flowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("step") & "'"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	msg = "Found existing step!"
end if

rsOra.close

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
	sqlOra = "select distinct cStepNextNo from flowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("selPrevStep") & "'"
	rsOra.open sqlOra, connOra
	
	if not rsOra.EOF then
		nextStep = rsOra("cStepNextNo")
	else
		nextStep = "00"
	end if
	
	rsOra.close
end if

'########## FIND APPROVER ##########
select case request.queryString("optApvType")
	case "1"
		apv = "'" & DEFAULT_APPROVER & "'"
		apvOrg = "null"
		
		if request.queryString("selPos") = "" then
			pos = "null"
		else
			pos = "'" & request.queryString("selPos") & "'"
		end if
		
	case "2"
		apv = "'" & DEFAULT_APPROVER & "'"
		
		if request.queryString("selOwnerPos") = "" then
			apvOrg = "null"
			pos = "null"
		else
			apvOrg = "'" & left(request.queryString("selOwnerPos"), inStr(request.queryString("selOwnerPos"), ",")-1) & "'"
			pos = "'" & mid(request.queryString("selOwnerPos"), inStrRev(request.queryString("selOwnerPos"), ",")+1, len(request.queryString("selOwnerPos"))) & "'"
		end if
		
	case "3"
		apvOrg = "null"
		pos = "null"
		
		if request.queryString("txtApv") = "" then
			apv = "'" & DEFAULT_APPROVER & "'"
		else
			apv = "'" & request.queryString("txtApv") & "'"
		end if	
end select
'########## END FIND APPROVER ##########

'Set apply all flag
if request.queryString("chkApvApply") = APPLY_ALL_APV then
	applyAll = APPLY_ALL_APV
else
	applyAll = APPLY_ALL_NONE
end if

if request.queryString("chkRejApply") = APPLY_ALL_REJ then
	applyAll = cInt(applyAll) + cInt(APPLY_ALL_APV)
end if

'Insert new step
sqlOra = "insert into flowMst (nFrmNo, vOrgNo, cYear, cStepNo, cStepNextNo, vUrl, vApvOrgNo, vPosNo, vApvNo, cStart, cType, cExtData, cApvType, cRejType, cApplyAll) " _
& "values (" & request.queryString("no") & ", '" & request.queryString("orgNo") & "', '" & request.queryString("y") & "', '" & request.queryString("selStep") & "', '" & nextStep & "', '" & request.queryString("txtUrl") & "', " & apvOrg & ", " & pos & ", " & apv & ", '" & start & "', '" & request.queryString("optApvType") & "', '" & request.queryString("txtExtData") & "', '" & request.queryString("optApvMethod") & "', '" & request.queryString("chkRejType") & "', '" & applyAll & "')"
'response.write sqlOra & "<br>"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if

'Update previous step
sqlOra = "update flowMst set cStepNextNo = '" & request.queryString("selStep") & "' where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("selPrevStep") & "'"
'response.write sqlOra & "<br>"
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