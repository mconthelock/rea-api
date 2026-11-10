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

'Check existing step
sqlOra = "select * from stepMst where cNo = '" & request.queryString("stepNo") & "'"
rsOra.open sqlOra, connOra

if rsOra.EOF then
	msg = "Step not found!"
end if

rsOra.close

'Find reference data
sqlOra = "select count(*) from flowMst where cStepNo = '" & request.queryString("stepNo") & "' or cStepNextNo = '" & request.queryString("stepNo") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if rsOra("count(*)") > 0 then
	msg = "Found reference data!"
end if

rsOra.close

if not isNumeric(request.queryString("stepNo")) then
	msg = "Invalid step number!"
end if

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Insert new step
sqlOra = "delete from stepMst where cNo = '" & request.queryString("stepNo") & "'"
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

response.redirect "../stepLst.asp"
%>