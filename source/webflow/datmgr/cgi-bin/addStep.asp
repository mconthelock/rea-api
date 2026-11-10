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
sqlOra = "select * from stepMst where lower(vName) = '" & lCase(oraFormatStr(left(trim(request.queryString("txtName")), 30))) & "'"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	msg = "Found existing step!"
end if

rsOra.close

if len(trim(request.queryString("txtName"))) = 0  then
	msg = "Step name not found!"
end if

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Get max cNo
sqlOra = "select max(cNo) from stepMst"
rsOra.open sqlOra, connOra

if isNull(rsOra("max(cNo)")) then
	no = "01"
else
	no = cInt(rsOra("max(cNo)")) + 1
	no = string(2 - len(no), "0") & no
end if

rsOra.close

'Insert new step
sqlOra = "insert into stepMst (cNo, vName, vDesc) values ('" & no & "', '" & oraFormatStr(left(trim(request.queryString("txtName")), 30)) & "', '" & oraFormatStr(left(trim(request.queryString("txtDesc")), 256)) & "')"
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