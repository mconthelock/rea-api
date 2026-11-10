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

'Find tree's child
sqlOra = "select count(*) from orgTree where vParent = '" & request.queryString("childNo") & "'"
rsOra.open sqlOra, connOra

if rsOra("count(*)") > 0 then
	msg = "Found references to this organize!"
end if

rsOra.close

if not isNumeric(request.queryString("orgNo")) then
	msg = "Invalid organization number!"
end if

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Insert new step
sqlOra = "delete from orgTree where vOrgNo = '" & request.queryString("childNo") & "' and vParent = '" & request.queryString("orgNo") & "'"
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