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

'Update C/C Master
sqlOra = "update  fxa_ccMst   " _
& "set ccCode = '" & request("ccCode") & "', ccDesc = '" & oraFormatStr(left(request("ccName"), 256)) & "' " _
& "where " _
& " ccCode= '" & request("ccCodeO") & "'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra
connOra.commitTrans

response.redirect "ccMst.asp"
%>