<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim empNo

validateAuthData

set rsOra = createObject("adodb.recordset")

connOra.beginTrans
showOraError connOra

'Delete current authority
sqlOra = "delete from rptUsrAuth " _
& "where " _
& "nFrmNo = '" & request.form("no") & "' " _
& "and vOrgNo = '" & request.form("orgNo") & "' " _
& "and cYear = '" & request.form("y") & "' and vEmpNo in ('" & replace(request.form("empList"), ", ", "', '") & "') "
response.write sqlOra & "<br>"
on error resume next
rsOra.open sqlOra, connOra
showOraError connOra

for each empNo in split(request.form("empList"), ", ")
	if request.form("optAuth" & empNo) <> "" then
		sqlOra = "insert into rptUsrAuth " _
		& "(nFrmNo, vOrgNo, cYear, vEmpNo, cAuthNo) " _
		& "values " _
		& "('" & request.form("no") & "', '" & request.form("orgNo") & "', '" & request.form("y") & "', '" & empNo & "', '" & mid(request.form("optAuth" & empNo), inStr(request.form("optAuth" & empNo), ",")+1, len(request.form("optAuth" & empNo))) & "')"
		'response.write sqlOra & "<br>"
		'response.end
		rsOra.open sqlOra, connOra
		showOraError connOra
	end if
next

connOra.commitTrans
showOraError connOra
on error goto 0

response.redirect "../rptAuthen.asp?selScope=" & server.urlEncode(request.form("scope")) & "&txtEmpNo=" & server.urlEncode(request.form("txtUid")) & "&txtName=" & server.urlEncode(request.form("txtUName")) & "&selSec=" & server.urlEncode(request.form("selSec")) & "&selDep=" & server.urlEncode(request.form("selDep")) & "&selDiv=" & server.urlEncode(request.form("selDiv")) & "&search=" & request.form("search")

function validateAuthData()
	if request("no") = "" or request("orgNo") = "" or request("y") = "" or request("empList") = "" then
		response.write "Report/user data not found!<br>Please choose report and set user's authority"
		response.end
	end if
end function
%>
