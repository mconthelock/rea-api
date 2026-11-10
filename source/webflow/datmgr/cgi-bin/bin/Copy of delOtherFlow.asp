<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no

set rsOra = createObject("adodb.recordset")

'Find reference data
if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Check root node
sqlOra = "select cStart from formFlowMst " _
& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
& "and cYear = '" & request.queryString("y") & "' and  cStepNo = '" & request.queryString("step") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

'If specific is root node then delete all flow
if rsOra("cStart") = "1" then	
	rsOra.close
	
	sqlOra = "delete from formFlowMst " _
	& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
	& "and cYear = '" & request.queryString("y") & "'"
	rsOra.open sqlOra, connOra

	on error resume next
	rsOra.open sqlOra, connOra
	on error goto 0
	
	if connOra.errors.count > 0 then
		response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
		response.end
	end if
else
	rsOra.close
	
	'Update previous step
	sqlOra = "update formFlowMst " _
	& "set cStepNextNo = '" & request.queryString("stepNext") & "' " _ 
	& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
	& "and cYear = '" & request.queryString("y") & "' and  cStepNo = '" & request.queryString("step") & "'"
	rsOra.open sqlOra, connOra
	
	on error resume next
	rsOra.open sqlOra, connOra
	on error goto 0
	
	if connOra.errors.count > 0 then
		response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
		response.end
	end if
	
	'Delete step
	sqlOra = "delete from formFlowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("step") & "' and cStepNextNo = '" & request.queryString("stepNext") & "'"
	'response.write sqlOra & "<br>"
	
	on error resume next
	rsOra.open sqlOra, connOra
	on error goto 0
	
	if connOra.errors.count > 0 then
		response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
		response.end
	end if
end if

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../frmLst.asp"
%>