<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim pos, apv, apvOrg
Dim applyAll

set rsOra = createObject("adodb.recordset")

connOra.beginTrans

'########## Update next node ##########
sqlOra = "update flowMst set cStepNextNo = '" & request.queryString("stepNext") & "' where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and   cStepNextNo = '" & request.queryString("step") & "'"
'response.write sqlOra & "<br>"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if
'########## END Update next node ##########

'########### Update current node ###########
'sqlOra = "update flowMst set cStepNextNo = (select distinct cStepNextNo from flowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and  cStepNo = '" & request.queryString("selPrevStep") & "') where cStepNo = '" & request.queryString("step") & "'"

if request.queryString("selPrevStep") <> "" then
	sqlOra = "update flowMst set " _
	& "cStepNextNo = (select distinct cStepNextNo from flowMst where nFrmNo = '" & request.queryString("no") & "' " _
	& "and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' " _
	& "and  cStepNo = '" & request.queryString("selPrevStep") & "') " _
	& "where " _
	& "nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
	& "and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("step") & "'"
	'response.write sqlOra & "<br>"

	on error resume next
	rsOra.open sqlOra, connOra
	on error goto 0
	
	if connOra.errors.count > 0 then
		response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
		response.end
	end if
end if
'########### END Update current node ###########

'Update node to be substituted
'sqlOra = "update flowMst set cStepNextNo = (select distinct cStepNextNo from flowMst where cStepNo = '" & request.queryStr & "') where cStepNo = (select cStepNextNo from flowMst where cStepNo = '" & request.queryString("selPrevStep") & "')"

'########### Update new previous node ###########
sqlOra = "update flowMst set cStepNextNo = '" & request.queryString("step") & "' where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and  cStepNo = '" & request.queryString("selPrevStep") & "'"
'response.write sqlOra & "<br>"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if
'########### END Update new previous node ###########

'########### Update current node data ###########
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
applyAll = APPLY_ALL_NONE

if request.queryString("chkApvApply") = APPLY_ALL_APV then
	applyAll = cInt(applyAll) + cint(APPLY_ALL_APV)
end if

if request.queryString("chkRejApply") = APPLY_ALL_REJ then
	applyAll = cInt(applyAll) + cInt(APPLY_ALL_REJ)
end if

sqlOra = "update flowMst set vApvOrgNo = " & apvOrg & ", vPosNo = " & pos & ", vApvNo = " & apv & ", vUrl = '" & request.queryString("txtUrl") & "', cType = '" & request.queryString("optApvType") & "', cExtData = '" & oraFormatStr(left(request.queryString("txtExtData"), 2)) & "', cApvType = '" & request.queryString("optApvMethod") & "', cRejType = '" & request.queryString("chkRejType") & "', cApplyAll = '" & applyAll & "' " _
& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("step") &  "'"
'response.write sqlOra & "<br>"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../editFrmFrm.asp?no=" & request.queryString("no") & "&orgNo=" & request.queryString("orgNo") & "&y=" & request.queryString("y")
%>