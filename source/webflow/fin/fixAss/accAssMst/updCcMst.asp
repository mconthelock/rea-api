<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra

if (trim(request("accCode")) = "") or (trim(request("accName")) = "") then
	response.write "Please enter C/C Code and C/C description!"
	response.end
end if

set rsOra = createObject("adodb.recordset")
connOra.beginTrans

'Update C/C Master
sqlOra = "update  fxa_accassMst " _
& "set accassCode = '" & request("accCode") & "', accassName = '" & oraFormatStr(left(request("accName"), 256)) & "' , vowner =  '" & request("vOwner") & "'" _
& "where " _
& " accassCode= '" & request("accCodeO") & "'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra
connOra.commitTrans

response.redirect "accAssMst.asp"
%>