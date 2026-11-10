<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no

set rsOra = createObject("adodb.recordset")

connOra.beginTrans

'Delete user authority
sqlOra = "delete from rptUsrAuth where nFrmNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' and vOrgNo = '" & request.queryString("orgNo") & "'"
'response.write sqlOra & "<br>"
on error resume next
rsOra.open sqlOra, connOra
showOraError connOra

'Delete report master
sqlOra = "delete from rptMst where nRepNo = '" & request.queryString("no") & "' and cYear = '" & request.queryString("y") & "' and vOrgNo = '" & request.queryString("orgNo") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra
showOraError connOra
on error goto 0

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../rptLst.asp"
%>