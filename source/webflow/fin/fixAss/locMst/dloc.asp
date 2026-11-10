<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->

<%
Dim rsOra, sqlOra

set rsOra = createObject("adodb.recordset")
connOra.beginTrans

'Delete location
sqlOra = "delete " _
& "from " _
& "fxa_locMst " _
& "where " _
& "locCode = '" & request("locCode") & "'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra


connOra.commitTrans

response.redirect "locMst.asp"
%>