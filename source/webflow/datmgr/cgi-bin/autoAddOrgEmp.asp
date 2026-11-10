<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, rsOra2, sqlOra
Dim orgField 'Field to be compare with orgNo, sSecCode, sDepCode and sDivCode
Dim orgClause

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

orgField = array("sSecCode", "sDepCode", "sDivCode")

'########## VERIFY DATA ##########
sqlOra = "select * from vOrgMst where vOrgNo = '" & request.queryString("orgNo") & "' and cType = '" & request.queryString("t") & "'"
rsOra.open sqlOra, connOra

if rsOra.EOF then
	response.write "Invalid data!"
	response.end
end if

rsOra.close
'########## END VERIFY DATA ##########

select case request.queryString("t")
	case "1"
		orgClause = "sSecCode = '" & request.queryString("orgNo") & "'"
	case "2"
		orgClause = "sSecCode = '00' and sDepCode = '" & request.queryString("orgNo") & "'"
	case "3"
		orgClause = "sSecCode = '00' and sDepCode = '00' and sDivCode = '" & request.queryString("orgNo") & "'"
end select

'Delete current employee
sqlOra = "delete from orgPos where vOrgNo = '" & request.queryString("orgNo") & "'"
rsOra.open sqlOra, connOra

connOra.beginTrans

'########## ADD EMPLOYEE ##########
sqlOra = "select sEmpNo, sPosCode from amec.aEmployee where " & orgClause & " and cStatus <> '0'"
rsOra.open sqlOra, connOra

do while not rsOra.EOF
	sqlOra = "insert into orgPos (vOrgNo, vEmpNo, vPosNo) values ('" & request.queryString("orgNo") & "', '" & rsOra("sEmpNo") & "', '" & rsOra("sPosCode") & "')"
	'response.write sqlOra & "<br>"
	
	on error resume next
	rsOra2.open sqlOra, connOra
	on error goto 0
	
	if connOra.errors.count > 0 then
		response.write connOra.errors(0).number & "<br>" & connOra.errors(0).description
		response.end
	end if
	
	rsOra.moveNext
loop

rsOra.close
'########## END ADD EMPLOYEE ##########

connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../editOrgFrm.asp?orgNo=" & request.queryString("orgNo")
%>