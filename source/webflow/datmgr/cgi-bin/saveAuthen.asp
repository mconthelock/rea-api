<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim empNo

set rsOra = createObject("adodb.recordset")

'Delete current authority
sqlOra = "delete from usrAuth " _
& "where " _
& "nFrmNo = '" & request.form("no") & "' " _
& "and vOrgNo = '" & request.form("orgNo") & "' " _
& "and cYear = '" & request.form("y") & "' and vEmpNo in ('" & replace(request.form("empList"), ", ", "', '") & "') "
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

for each empNo in split(request.form("empList"), ", ")
	if request.form("optAuth" & empNo) <> "" then
		sqlOra = "insert into usrAuth " _
		& "(nFrmNo, vOrgNo, cYear, vEmpNo, cAuthNo) " _
		& "values " _
		& "('" & request.form("no") & "', '" & request.form("orgNo") & "', '" & request.form("y") & "', '" & empNo & "', '" & mid(request.form("optAuth" & empNo), inStr(request.form("optAuth" & empNo), ",")+1, len(request.form("optAuth" & empNo))) & "')"
		'response.write sqlOra & "<br>"
		rsOra.open sqlOra, connOra
	end if
next

response.redirect "../formAuthen.asp?no=" & request.form("no") & "&orgNo=" & request.form("orgNo") & "&y=" & request.form("y") & "&txtUid=" & server.urlEncode(request.form("txtUid")) & "&txtUName=" & server.urlEncode(request.form("txtUName")) & "&selSec=" & server.urlEncode(request.form("selSec")) & "&selDep=" & server.urlEncode(request.form("selDep")) & "&selDiv=" & server.urlEncode(request.form("selDiv")) & "&search=" & request.form("search")
%>