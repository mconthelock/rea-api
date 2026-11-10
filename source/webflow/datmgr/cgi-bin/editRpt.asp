<%
Option Explicit
%>

<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no, lifeTime

validateReportData

set rsOra = createObject("adodb.recordset")

connOra.beginTrans

'Update report
sqlOra = "update " _
& "rptMst " _
& "set vName = '" & oraFormatStr(left(trim(request("txtName")), 120)) & "' " _
& ", vDesc = '" & oraFormatStr(left(trim(request("txtDesc")), 256)) & "' " _
& ", vUrl =  '" & oraFormatStr(left(trim(request("txtUrl")), 1024)) & "' " _
& ", cStatus = '" & request("cVisible") & "' " _
& "where " _
& "nRepNo = '" & request("no") & "' and vOrgNo = '" & request("orgNo") & "' " _
& "and cYear = '" & request("y") & "'"
'response.write sqlOra & "<br>"

on error resume next
rsOra.open sqlOra, connOra
showOraError connOra
on error goto 0

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../rptLst.asp"

function validateReportData()
	if trim(request("txtName")) = "" then
		response.write "Report's name not found!"
		response.end
	end if
	
	if trim(request("txtUrl")) = "" then
		response.write "Report's URL not found!"
		response.end
	end if
end function
%>
