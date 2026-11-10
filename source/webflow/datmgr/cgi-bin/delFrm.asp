<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no

set rsOra = createObject("adodb.recordset")

'Find flow reference data
sqlOra = "select count(*) from form where " _
& "nFrmNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' " _
& "and vOrgNo = '" & request.queryString("orgNo") & "'"
rsOra.open sqlOra, connOra

if rsOra("count(*)") > 0 then
	msg = "Found reference data!"
end if

rsOra.close

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Delete user authority
sqlOra = "delete from usrAuth where nFrmNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' and vOrgNo = '" & request.queryString("orgNo") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

'Delete acknowledge by email data
sqlOra = "delete from emailAck where nFrmNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' and vOrgNo = '" & request.queryString("orgNo") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

'Delete representative
sqlOra = "delete from rep where nFrmNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' and vOrgNo = '" & request.queryString("orgNo") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

'Delete flow step master
sqlOra = "delete from flowMst where nFrmNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' and vOrgNo = '" & request.queryString("orgNo") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

'Delete form master
sqlOra = "delete from formMst where nNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' and vOrgNo = '" & request.queryString("orgNo") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if connOra.errors.count > 0 then
	response.write "Unable to delete this form, references data found!"
	response.end
end if

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../frmLst.asp"
%>