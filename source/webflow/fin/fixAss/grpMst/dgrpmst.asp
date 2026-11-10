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
& " fxa_grpMst  " _
& "where " _
& "grpCode = '" & request("grpCode") & "'"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra


connOra.commitTrans

response.redirect "grpMst.asp"
%>