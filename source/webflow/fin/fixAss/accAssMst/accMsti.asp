<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra

if (trim(request("ccCode")) = "") or (trim(request("ccName")) = "") then
	response.write "Please enter ASSET CODE and ASSET NAME!"
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
sqlOra = "insert into fxa_accassMst " _
& "(accassCode, accassName, vOwner) " _
& "values " _
& "('" & nextrunno & "', '" & oraFormatStr(left(request("ccName"), 256)) & "','" & request("txtvowner") & "') "
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra

connOra.commitTrans

response.redirect "accMst.asp"
%>