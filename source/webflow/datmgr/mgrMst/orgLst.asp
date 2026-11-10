<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/misc.inc"-->
<!--#include file="../inc/datMgr.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim i, aCode(1000)

set rsOra = createObject("adodb.recordset")

'sqlOra = "select distinct 10*(lv-1) as lv, cStart, vOrgMst.* from vOrgTree, vOrgMst " _
sqlOra = "select distinct lv, cStart, vOrgMst.* from vOrgTree, vOrgMst " _
& "where " _
& "vOrgMst.vOrgNo = vOrgTree.vOrgNo order by vOrgMst.vOrgNo"
'& "order by lv, vOrgMst.vOrgNo"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra
%>
<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Organization list</title>
</head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">

<%showPageLoading%>

<form name="frmMain" method="get" action="">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
            <td align="right" valign="bottom"><img src="../../imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="../../imgs/border/book/t.gif" width="100%" align="center"><img src="../../imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../../imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="../../imgs/border/book/l.gif" align="right">&nbsp;</td>
            
      <td bgcolor="#AACCEE" align="center"><b>Organization Master</b></td>
            <td background="../../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
            <td bgcolor="#0000FF"> 
				
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow"> 
					<!--<td nowrap width="1%" align="center"><b>No</b></td>-->
					
            <td nowrap align="center"><b>Organize</b></td>
					
            <td nowrap align="center"><b>Description</b></td>
					
            <td nowrap align="center">&nbsp;</td>
					
            <td nowrap align="center">&nbsp;</td>
				  </tr>
				  <%
					if not rsOra.EOF then
						i = 1
						
						do while not rsOra.EOF
							buildTree "<font color=""#0000FF"">" & rsOra("vOrgNo") & "</font>", rsOra("lv")-1, 1
							if 1 = 2 then
					%>
					
          <tr bgcolor="<%
					if i mod 2 = 0 then
						response.write "#FFCC00"
					else
						response.write "#FFFFCC"
					end if
					%>"> 
            	<!--<td align="center" width="1%" nowrap><font color="#000000"><%if rsOra("cStart") = "1" then response.write i else response.write "&nbsp;" end if%></font></td>-->
            <td nowrap><font color="#000000"><img src="../../imgs/spacer.gif" height="1" width="<%=rsOra("lv")%>">(<%=rsOra("vOrgNo")%>)&nbsp;<%=rsOra("vName")%></font></td>
            <td nowrap><font color="#000000"><%=rsOra("vDesc")%></font></td>
            <td align="center" nowrap> 
              <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editOrgFrm.asp?orgNo=<%=rsOra("vOrgNo")%>&orgName=<%=server.urlEncode(rsOra("vName") & "")%>&orgDesc=<%=server.urlEncode(rsOra("vDesc") & "")%>'">
            </td>
            <td align="center" nowrap> 
              <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete tree?\n\n\'<%=rsOra("vName") & ": " & replace(replace(rsOra("vDesc"), "'", "\'"), QUOTE, "\" & QUOTE)%>\'')) {window.location.href = 'cgi-bin/delOrg.asp?orgNo=<%=rsOra("vOrgNo")%>'}">
            </td>
          </tr>
          <%
		  end if
							i = i + 1
							rsOra.moveNext
						loop
					else
					%>
          <tr> 
            <td align="center" colspan="4" bgcolor="#FFFFFF"><br>
              <font class="alertMsg"><b>Data not found!</b></font><br>
              &nbsp;</td>
          </tr>
          <%
					end if
					
					rsOra.close
					%>
        </table>
            </td>
            <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="../../imgs/border/book/bg.gif">&nbsp;</td>
            <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif"></td>
            <td align="center" background="../../imgs/border/book/bg.gif"> 
                <input type="button" value="Add organization..." onClick="javascript: window.location.href = 'addOrgFrm.asp'">
                <input type="button" value="Back" onClick="javascript: window.location.href = 'index.asp'">
            </td>
            <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="../../imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="../../imgs/border/book/b.gif"><img src="../../imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../../imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
</form>
<!--#include file="../../inc/copyRight.inc"-->

<%hidePageLoading%>

</body>
</html>


<%
set rsOra = nothing
%>