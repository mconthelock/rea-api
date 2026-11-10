<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/userUtil.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim msg, emp, pos

set rsOra = createObject("adodb.recordset")

'Check existing step
if 1=2 then
sqlOra = "select vOrgNo from orgPos where vEmpNo in ('" & replace(request.queryString("chkEmp"), ", ", "', '") & "') and vOrgNo = '" & request.queryString("orgNo") & "' and vPosNo in ('" & replace(replace(request.queryString("selPos"), ", ", "', '"), "&&", "', '") & "')"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	msg = "Found existing employee in specific position!"
end if

rsOra.close
end if

'Delete current selected employee
sqlOra = "delete from orgPos where vOrgNo = '" & request.queryString("orgNo") & "' and vEmpNo in ('" & replace(request.queryString("chkEmp"), ", ", "', '") & "') and vPosNo in ('" & replace(replace(request.queryString("selPos"), ", ", "', '"), "&&", "', '") & "')"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if msg <> "" then
	response.write msg
	response.end
end if

connOra.beginTrans

'Insert new step
for each emp in split(request.queryString("chkEmp"), ", ")
	for each pos in split(request.queryString("selPos"), ", ")
		if left(pos, inStr(pos, "&&") - 1) = emp then
			'Check working user
			if chkWorkingUser(emp) then
				response.write "Employee not found!"
				response.end
			end if
			
			'Insert user
			sqlOra = "insert into orgPos (vEmpNo, vOrgNo, vPosNo) values ('" & emp & "', '" & request.queryString("orgNo") & "', '" &  mid(pos, inStr(pos, "&&") + 2, len(pos)) & "')"
			'response.write sqlOra & "<br>"
			
			on error resume next
			rsOra.open sqlOra, connOra
			on error goto 0
			
			if connOra.errors.count > 0 then
				response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
				response.end
			end if
		end if
	next
next

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../editOrgFrm.asp?orgNo=" & request.queryString("orgNo")
%>