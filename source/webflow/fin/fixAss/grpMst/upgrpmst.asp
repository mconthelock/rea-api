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

'Update location
sqlOra = "update  fxa_grpMst  " _
& "set grpCode= '" & request("grpCode") & "', grpdesc =  '" & oraFormatStr(left(request("grpName"), 256)) & "' " _
& "where " _
& "grpCode= '" & request("grpCodeO") & "'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra
connOra.commitTrans

response.redirect "grpMst.asp"
%>