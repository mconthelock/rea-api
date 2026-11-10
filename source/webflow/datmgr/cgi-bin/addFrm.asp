<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/date.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, no, runNo, start

set rsOra = createObject("adodb.recordset")

'Validate data
if trim(request.queryString("txtName")) = "" then
	response.write "Form name not found!"
	response.end
end if

'Check existing form
sqlOra = "select nNo from formMst where vOrgNo = '" & request.queryString("selOrg") & "' and lower(vName) = '" & oraFormatStr(lCase(request.queryString("txtName"))) & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	msg = "Found existing form!"
end if

rsOra.close

if msg <> "" then
	response.write msg
	response.end
end if

'Get form number
sqlOra = "select max(nNo) from formMst where cYear = '" & right(year(date), 2) & "'"
rsOra.open sqlOra, connOra

if isNull(rsOra("max(nNo)")) then
	no = 1
else
	no = cInt(rsOra("max(nNo)")) + 1
end if

rsOra.close

'Get form running number
sqlOra = "select max(nRunNo) from formMst where cYear = '" & right(year(date), 2) & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if isNull(rsOra("max(nRunNo)")) then
	runNo = 1
else
	runNo = cInt(rsOra("max(nRunNo)")) + 1
end if

rsOra.close

connOra.beginTrans

'Insert new step
sqlOra = "insert into formMst (nNo, cYear, nRunNo, vName, vDesc, vOrgNo, dCreDate, cCreTime) " _
& "values " _
& "(" & no & ", '" & right(year(date), 2) & "', '" & runNo & "', '" & oraFormatStr(left(trim(request.queryString("txtName")), 60)) & "', '" & oraFormatStr(left(trim(request.queryString("txtDesc")), 256)) & "', '" & request.queryString("selOrg") & "', to_date('" & toShortDate(date, "DD/MM/YYYY") & "', 'DD/MM/YYYY'), '" & toFullTime(time) & "')"
'response.write sqlOra & "<br>"
'response.end

on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write "Can't create form! try again." 'connOra.errors(0).number & "<br>" & connOra.errors(0).description
	response.end
end if

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../frmLst.asp"
%>