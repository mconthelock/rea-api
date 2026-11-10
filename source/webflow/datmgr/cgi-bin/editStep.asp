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
sqlOra = "select * from stepMst where cNo = '" & request.queryString("cNo") & "'"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	msg = "Step not found!"
end if

rsOra.close

if len(trim(request.queryString("txtName"))) = 0  then
	msg = "Step name not found!"
end if

if not isNumeric(request.queryString("stepNo")) then
	msg = "Invalid step number!"
end if

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Insert new step
sqlOra = "update stepMst set vName = '" & oraFormatStr(left(trim(request.queryString("txtName")), 30)) & "', vDesc = '" & oraFormatStr(left(trim(request.queryString("txtDesc")), 256)) & "' where cNo = '" & request.queryString("stepNo") & "'"
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