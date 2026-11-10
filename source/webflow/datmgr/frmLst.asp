<%Option Explicit%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i

set rsOra = createObject("adodb.recordset")

sqlOra = "select to_char(nNo) as nNo, vOrgNo, cYear, dCreDate, cCreTime, vName, vAName, vDesc, vAuthPage from formMst " _
& "order by vName"
if request("ownerNo") = "0504" then
	sqlOra = "select to_char(nNo) as nNo, vOrgNo, cYear, dCreDate, cCreTime, vName, vAName, vDesc, vAuthPage from formMst " _
	& "where vOrgno like '" & request("ownerNo") & "%' or vOrgno like '0514%' " _
	& "order by vName"
elseif request("ownerNo") = "0503" then
	sqlOra = "select to_char(nNo) as nNo, vOrgNo, cYear, dCreDate, cCreTime, vName, vAName, vDesc, vAuthPage from formMst " _
	& "where vOrgno like '" & request("ownerNo") & "%' or vOrgno like '00%' " _
	& "order by vName"	
else
	sqlOra = "select to_char(nNo) as nNo, vOrgNo, cYear, dCreDate, cCreTime, vName, vAName, vDesc, vAuthPage from formMst " _
	& "where vOrgno like '" & request("ownerNo") & "%' " _
	& "order by vName"
end if
'response.write sqlOra & "<br>"
'response.end
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
            <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Form 
                list</font></b></td>
            <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow" align="center"> 
            <td nowrap width="1%"><b><font color="#000000">No</font></b></td>
            <td nowrap><b><font color="#000000">Alias</font></b></td>
            <td nowrap><b><font color="#000000">Name</font></b></td>
            <td nowrap><b><font color="#000000">Description</font></b></td>
            <td nowrap align="center" width="1%"><b><font color="#000000">Org 
              No </font></b></td>
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
            <td nowrap><font color="#000000"><%=rsOra("vAName")%></font></td>
            <td nowrap width="100%"><font color="#000000"><%=rsOra("vName")%></font></td>
            <td nowrap width="100%"><font color="#000000"><%=showData(rsOra("vDesc"))%></font></td>
            <td nowrap width="100%" align="center">&nbsp;</td>
            <!--<td align="center" nowrap width="1%"><%
						if isNull(rsOra("vAuthPage")) then
							response.write "&nbsp;"
						else
						%>
						<input type="button" value="Authority..." onClick="javascript: window.location.href = '<%=rsOra("vAuthPage")%>?<%="no=" & rsOra("nNo") & "&orgNo=" & rsOra("vOrgNo") & "y=" & rsOra("cYear")%>'"><%end if%></td>-->
            <td align="center" nowrap width="1%"> 
              <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editFrmFrm.asp?no=<%=rsOra("nNo")%>&orgNo=<%=rsOra("vOrgNo")%>&y=<%=rsOra("cYear")%>&ownerNo=<%=request("ownerNo")%>'">
            </td>
            <td align="center" nowrap width="1%"> 
              <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this form?\n\n\Form:<%=rsOra("nNo") & "\nName: " & replace(rsOra("vName") & "", "'", "\'") & "\nDesc: " & replace(replace(rsOra("vDesc") & "", "'", "\'"), QUOTE, "\" & QUOTE)%>')) {window.location.href = 'cgi-bin/delFrm.asp?no=<%=rsOra("nNo")%>&orgNo=<%=rsOra("vOrgNo")%>&y=<%=rsOra("cYear")%>'}">
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
                <input type="button" value="Add form..." onClick="javascript: window.location.href = 'addFrmFrm.asp'">
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