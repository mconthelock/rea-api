<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->

<%

Dim rsOra, sqlOra

if (trim(request("grpCode")) = "" ) or (trim(request("grpName")) = "" ) then
	response.write "Please enter group code or group description!"
	response.end
end if

set rsOra = createObject("adodb.recordset")
connOra.beginTrans

'sqlOra = "select max(nvl(locCode, 0)) + 1 as nextrunno from fxa_locMst"
'rsOra.open sqlOra, connOra

'if rsOra("nextrunno") > 99999 then
'	response.write "!"
'	response.end
'else
'	nextrunno = rsOra("nextrunno")
'end if

'rsOra.close

'Update Location Master
sqlOra = "insert into fxa_grpMst " _
& "(grpCode, grpdesc) " _
& "values " _
& "('" & request("grpCode") & "', '" & oraFormatStr(left(request("grpName"), 256)) & "') "
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra

connOra.commitTrans

response.redirect "grpMst.asp"
%>