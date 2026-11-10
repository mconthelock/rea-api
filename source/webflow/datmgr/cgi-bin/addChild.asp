<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no, start, child

set rsOra = createObject("adodb.recordset")

'Check existing step
if inStr(request.queryString("chkChild"), request.queryString("orgNo")) > 0 then
	msg = "Found recursive setting!"
end if

sqlOra = "select vOrgNo from orgTree where vOrgNo in ('" & replace(request.queryString("chkChild"), ", ", "', '") & "') and vParent = '" & request.queryString("orgNo") & "'"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	msg = "Found existing child!"
end if

rsOra.close

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Insert new step
for each child in split(request.queryString("chkChild"), ", ")
	sqlOra = "insert into orgTree (vOrgNo, vParent, cStart) values ('" & trim(child) & "', '" & request.queryString("orgNo") & "', '0')"
	'response.write sqlOra & "<br>"
	
	on error resume next
	rsOra.open sqlOra, connOra
	on error goto 0
	
	if connOra.errors.count > 0 then
		response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
		response.end
	end if
next

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../editOrgFrm.asp?orgNo=" & request.queryString("orgNo")
%>