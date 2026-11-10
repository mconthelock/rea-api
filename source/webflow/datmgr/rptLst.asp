<%Option Explicit%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i

set rsOra = createObject("adodb.recordset")

sqlOra = "select to_char(nRepNo) as nRepNo, vOrgNo, cYear, vName, vDesc, vUrl, cStatus from rptMst " _
& "order by vName"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Form list</title></head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">
<form name="frmMain" method="get" action="">
    
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
            <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr>
            <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
            
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Report list</font></b></td>
            <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow" align="center"> 
            <td nowrap width="1%"><b><font color="#000000">No</font></b></td>
            <td nowrap><b><font color="#000000">Name</font></b></td>
            <td nowrap><b><font color="#000000">Description</font></b></td>
            <td nowrap width="1%"><font color="#000000"><b>URL</b></font></td>
            <td nowrap width="1%"><font color="#000000"><b>STATUS</b></font></td>
            <!--<td nowrap width="1%"><b><font color="#000000">Authority</font></b></td>-->
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
            <td align="center" width="1%" nowrap><font color="#000000"><%=i%></font></td>
            <td nowrap width="100%"><font color="#000000"><%=rsOra("vName")%></font></td>
            <td nowrap width="100%"><font color="#000000"><%=showData(rsOra("vDesc"))%></font></td>
            <td nowrap width="1%"><font color="#000000"><%=showData(rsOra("vUrl"))%></font></td>
            <td align="center" nowrap width="1%"><font color="#000000"><%if rsOra("cStatus") = RPT_STS_SHOW then response.write "visible" else response.write "hidden" end if%></font></td>
            <td align="center" nowrap width="1%"> 
              <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editRptFrm.asp?no=<%=rsOra("nRepNo")%>&orgNo=<%=rsOra("vOrgNo")%>&y=<%=rsOra("cYear")%>'">
            </td>
            <td align="center" nowrap width="1%"> 
              <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this report?\n\n\Report:<%=rsOra("nRepNo") & "\nName: " & replace(rsOra("vName") & "", "'", "\'") & "\nDesc: " & replace(replace(rsOra("vDesc") & "", "'", "\'"), QUOTE, "\" & QUOTE)%>')) {window.location.href = 'cgi-bin/delRpt.asp?no=<%=rsOra("nRepNo")%>&orgNo=<%=rsOra("vOrgNo")%>&y=<%=rsOra("cYear")%>'}">
            </td>
          </tr>
          <%
							i = i + 1
							rsOra.moveNext
						loop
					else
					%>
          <tr> 
            <td align="center" colspan="8" bgcolor="#FFFFFF"><br>
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
                
        <input type="button" value="Add report..." onClick="javascript: window.location.href = 'addRptFrm.asp'">
                <input type="button" value="Back" onClick="javascript: window.location.href = 'index.asp'">
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