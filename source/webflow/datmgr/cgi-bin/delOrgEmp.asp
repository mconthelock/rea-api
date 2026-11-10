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

connOra.beginTrans

'Insert new step
sqlOra = "delete from orgPos where vEmpNo = '" & request.queryString("empNo") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and vPosNo = '" & request.queryString("posNo") & "'"
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

response.redirect "../editOrgFrm.asp?orgNo=" & request.queryString("orgNo")
%>