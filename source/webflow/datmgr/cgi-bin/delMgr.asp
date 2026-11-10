<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include virtual="/sys/inc/numeric.inc"-->

<%
Dim rsOra, sqlOra

set rsOra = createObject("adodb.recordset")

connOra.beginTrans

'Delete manager
sqlOra = "delete from sequenceOrg " _
& "where " _
& "empNo " & translateFieldCompareClause(oraFormatStr(request("e"))) & " and vOrgNo " & translateFieldCompareClause(request("o")) & " and sPosCode " & translateFieldCompareClause(request("p")) & " " _
& "and cCo " & translateFieldCompareClause(request("c")) & " and headNo " & translateFieldCompareClause(oraFormatStr(request("h"))) & " and vOrgNo1 " & translateFieldCompareClause(request("o1")) & " " _
& "and sPosCode1 " & translateFieldCompareClause(request("p1")) & " and cCo1 " & translateFieldCompareClause(request("c1")) & " "
'response.write sqlOra & "<BR>"
rsOra.open sqlOra, connOra

connOra.commitTrans

set rsOra = nothing
response.redirect "../mgrList.asp" & "?selDiv=" & request("selDiv") & "&selDep=" & request("selDep") & "&selSec=" & request("selSec")
%>