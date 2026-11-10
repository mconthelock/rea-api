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

'Insert new step
sqlOra = "update formFlowMst set vPosNo = " & pos & ", vApvNo = " & apv & " where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStepNo = '" & request.queryString("step") &  "'"
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