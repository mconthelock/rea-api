<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim pc
Dim empNo, i

set rsOra = createObject("adodb.recordset")
connOra.beginTrans

'sqlOra = "delete from evapostree where vparent = '" & request("p") & "' "
'rsOra.open sqlOra, connOra
empNo = split(request("c"),",")
i=0
for each pc in split(replace(request("c"), " ", ""), ",")
'	sqlOra = "insert into evapostree (vparent, vposcode) values ('" & request("p") & "', '" & pc & "') "
	sqlOra = "delete from sequenceOrg where empno='" & trim(empNo(i)) & "'"
	'response.write sqlOra & "<br>"
	'response.end
	rsOra.open sqlOra, connOra

	sqlOra = "select sposcode from amec.aemployee where sempno='" &  trim(empNo(i)) & "'"
	rsOra.open sqlOra, connOra

	sqlOra = "insert into sequenceOrg(empno, sposcode, vorgno, cco, headno, sposcode1, vorgno1, cco1)" _
	& " values('" & trim(empNo(i)) & "','" & rsOra("sposcode") & "','" & request("orgno") & "','0','" & request("headno") & "','" & request("posNo") & "','" & request("orgno") & "','0')"
	response.write sqlOra & "<br>"
	rsOra.close
	rsOra.open sqlOra, connOra
	i=i+1
next
response.end
connOra.commitTrans
connOra.close

set rsOra = nothing

response.redirect "../posl.asp?orgNo=" & request("orgno")
%>