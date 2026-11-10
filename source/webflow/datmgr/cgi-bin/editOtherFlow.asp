<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, pos, apv

set rsOra = createObject("adodb.recordset")

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

if request.queryString("optApv") = "2" then
	pos = "null"
	apv = "'" & request.queryString("txtApv") & "'"
else
	pos = "'" & request.queryString("selPos") & "'"
	apv = "null"
end if

'########## Update current previous node ##########
sqlOra = "update orgFlowMst set cStepNextNo = '" & request.queryString("stepNext") & "' where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNextNo = '" & request.queryString("step") & "'"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if
'########## END Update current previous node ##########

'########### Update current node ###########
sqlOra = "update orgFlowMst set cStepNextNo = (select distinct cStepNextNo from orgFlowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("selPrevStep") & "') where cStepNo = '" & request.queryString("step") & "'"
'response.write sqlOra & "<br>"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if
'########### END Update current node ###########

'Update node to be substituted
'sqlOra = "update orgFlowMst set cStepNextNo = (select distinct cStepNextNo from orgFlowMst where cStepNo = '" & request.queryStr & "') where cStepNo = (select cStepNextNo from orgFlowMst where cStepNo = '" & request.queryString("selPrevStep") & "')"

'########### Update new previous node ###########
sqlOra = "update orgFlowMst set cStepNextNo = '" & request.queryString("step") & "' where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("selPrevStep") & "'"

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if
'########### END Update new previous node ###########

'########### Update current node data ###########
sqlOra = "update orgFlowMst set vPosNo = " & pos & ", vUrl = '" & request.queryString("txtUrl") & "' where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("step") &  "'"
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