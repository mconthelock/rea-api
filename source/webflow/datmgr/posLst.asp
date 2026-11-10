<%Option Explicit%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i

set rsOra = createObject("adodb.recordset")

sqlOra = "select * from (select level as lv, posTree.* from posTree start with cStart = '1' connect by vParent = prior vOrgNo), amec.pPosition where sPosCode = vOrgNo order by lv"
rsOra.open sqlOra, connOra
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Position tree</title></head>
<body>
<form name="frmMain" method="get" action="">
    <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
            <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr>
            <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
            <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Position 
                tree</font></b></td>
            <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            <td bgcolor="#0000FF"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="2">
                    <tr class="hdrRow" align="center"> 
                        <td nowrap width="1%"><b><font color="#000000">No</font></b></td>
                        <td nowrap width="1%"><b><font color="#000000">Level</font></b></td>
                        <td nowrap width="1%"><b><font color="#000000">Pos No</font></b></td>
                        <td nowrap><b><font color="#000000">Name</font></b></td>
                        <td nowrap><b><font color="#000000">Description</font></b></td>
                        <td nowrap width="1%">&nbsp;</td>
                        <td nowrap width="1%">&nbsp;</td>
                    </tr>
                    <%
					if not rsOra.EOF then
						i = 1
						
						do while not rsOra.EOF
					%>
                    <tr bgcolor="<%
					if i mod 2 = 0 then
						response.write "#FFCC00"
					else
						response.write "#FFFFCC"
					end if
					%>"> 
                        <td align="center" width="1%" nowrap><%=i%></td>
                        <td align="center" width="1%" nowrap><%=rsOra("lv")%></td>
                        <td align="center" width="1%" nowrap><%=rsOra("vPosNo")%></td>
                        <td nowrap><%=rsOra("vName")%></td>
                        <td nowrap width="100%"><%=rsOra("vDesc")%></td>
                        <td align="center" nowrap width="1%"> 
                            <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editStepFrm.asp?stepNo=<%=rsOra("cNo")%>&sname=<%=server.urlEncode(rsOra("vName"))%>&sdesc=<%=server.urlEncode(rsOra("vDesc") & "")%>'">
                        </td>
                        <td align="center" nowrap width="1%"> 
                            <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete step?\n\n\'<%=rsOra("cNo") & ": " & replace(replace(rsOra("vName"), "'", "\'"), QUOTE, "\" & QUOTE)%>\'')) {window.location.href = 'cgi-bin/delStep.asp?stepNo=<%=rsOra("cNo")%>'}">
                        </td>
                    </tr>
                    <%
							i = i + 1
							rsOra.moveNext
						loop
					else
					%>
                    <tr> 
                        <td align="center" colspan="7" bgcolor="#FFFFFF"><br>
                            <font class="alertMsg"><b>Data not found!</b></font><br>
                            &nbsp;</td>
                    </tr>
                    <%
					end if
					
					rsOra.close
					%>
                </table>
            </td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/bg.gif">&nbsp;</td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif"></td>
            <td align="center" background="../imgs/border/book/bg.gif"> 
                <input type="button" value="Add position..." onClick="javascript: window.location.href = 'addStepFrm.asp'">
                <input type="button" value="Back" onClick="javascript: window.location.href = '../index.html'">
            </td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="../imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
</form>
<!--#include file="../inc/copyRight.inc"-->
</body>
</html>

<%
set rsOra = nothing
%>