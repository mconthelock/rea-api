<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra

if (trim(request("ccCode")) = "") or (trim(request("ccName")) = "") then
	response.write "Please enter C/C Code and C/C description!"
	response.end
end if

set rsOra = createObject("adodb.recordset")
connOra.beginTrans

'sqlOra = "select max(nvl(ccCode, 0)) + 1 as nextrunno from fxa_ccMst"
'rsOra.open sqlOra, connOra

'if rsOra("nextrunno") > 99999 then
'	response.write "!"
'	response.end
'else
'	nextrunno = rsOra("nextrunno")
'end if

'rsOra.close
nextrunno = request("ccCode")
'Update Location Master
sqlOra = "insert into fxa_ccMst " _
& "(ccCode, ccDesc) " _
& "values " _
& "('" & nextrunno & "', '" & oraFormatStr(left(request("ccName"), 50)) & "') "
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra

connOra.commitTrans

response.redirect "ccMst.asp"
%>