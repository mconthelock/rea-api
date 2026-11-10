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

'Update Location Master
sqlOra = "insert into fxa_locMst " _
& "(locCode, locName, vorgno, sposcode) " _
& "values " _
& "('" & request("locCode") & "', '" & oraFormatStr(left(request("locName"), 50)) & "','" & request("selPos") & "', '" & request("selOrg") & "') "
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra

connOra.commitTrans

response.redirect "locMst.asp"
%>