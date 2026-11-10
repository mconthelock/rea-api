<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no, lifeTime

set rsOra = createObject("adodb.recordset")

'Check existing step
sqlOra = "select * " _
& "from " _
& "formMst " _
& "where " _
& "nNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
& "and cYear = '" & request.queryString("y") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if rsOra.EOF then
	msg = "Form master not found!"
end if

rsOra.close

if len(trim(request.queryString("txtName"))) = 0  then
	msg = "Form name not found!"
end if

if isNumeric(request.queryString("txtLifeTime")) then
	lifeTime = request.queryString("txtLifeTime")
else
	lifeTime = 6
end if

if lifeTime <= 0 then
	msg = "Life time must greater then zero!"
end if

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Insert new step
sqlOra = "update " _
& "formMst " _
& "set vName = '" & oraFormatStr(left(trim(request.queryString("txtName")), 60)) & "' " _
& ", vAName = '" & oraFormatStr(left(trim(request.queryString("txtAName")), 20)) & "' " _
& ", vDesc = '" & oraFormatStr(left(trim(request.queryString("txtDesc")), 256)) & "' " _
& ", vAuthPage =  '" & request.queryString("txtAuthPage") & "', vFormPage = '" & request.queryString("txtFormPage") & "' " _
& ", cStatus = '" & request.queryString("cVisible") & "', nLifeTime = '" & lifeTime & "' " _
& "where " _
& "nNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
& "and cYear = '" & request.queryString("y") & "'"
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