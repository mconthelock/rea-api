<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no, start

set rsOra = createObject("adodb.recordset")

'Check existing step
sqlOra = "select vOrgNo from orgTree where vOrgNo = '" & request.queryString("selOrg") & "'"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	msg = "Found existing organization!"
end if

rsOra.close

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Insert new step
sqlOra = "insert into orgTree (vOrgNo, vParent, cStart) values ('" & request.queryString("selOrg") & "', '00000', '1')"
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

response.redirect "../orgLst.asp"
%>