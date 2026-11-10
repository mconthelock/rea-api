<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra

if (trim(request("locCode")) = "") or (trim(request("locName")) = "") then
	response.write "Please enter location description!"
	response.end
end if

set rsOra = createObject("adodb.recordset")
connOra.beginTrans

'Update location
sqlOra = "update fxa_locMst  " _
& " set locCode= '" & trim(request("locCode")) & "', locName = '" & oraFormatStr(left(request("locName"), 256)) & "' " _
& " , vorgno = '" & request("selOrg") & "', sposcode = '" & request("selPos") & "'" _
& "where " _
& "locCode= '" & request("locCodeO") & "'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra
connOra.commitTrans

response.redirect "locMst.asp"
%>