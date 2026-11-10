<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->

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
sqlOra = "select cStart from flowMst " _
& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
& "and cYear = '" & request.queryString("y") & "' and  cStepNo = '" & request.queryString("step") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

'If specific is root node then delete all flow
if rsOra("cStart") = "1" then	
	rsOra.close
	
	sqlOra = "delete from flowMst " _
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

	deleteFlowStep connOra, request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), "flowMst", request.queryString("step"), request.queryString("stepNext")
end if

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../editFrmFrm.asp?no=" & request.queryString("no") & "&orgNo=" & request.queryString("orgNo") & "&y=" & request.queryString("y")
%>