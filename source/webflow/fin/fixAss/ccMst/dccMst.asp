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
& "fxa_ccMst " _
& "where " _
& "ccCode = '" & request("ccCode") & "'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra


connOra.commitTrans

response.redirect "ccMst.asp"
%>