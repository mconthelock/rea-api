<!--#include file="../../inc/connUtil.inc"--> 
<!--#include file="../../inc/db.inc"--> 
<%
Dim rsDes , sql , assno , desc 
set rsDes = createObject("adodb.recordset")
assno = trim(request.querystring("a"))
sql = "select ASSDESC , UNIT from fxa_faMst where ASSETNO = '"&assno&"' and statusass <> 'R'"

rsDes.open sql, connOra
if NOT rsDes.EOF then
	desc = rsDes("ASSDESC")&"|"&rsDes("UNIT")
else
	desc = "|"
end if
response.write desc

'response.write "jjj"&connOra&"jjj"
%>