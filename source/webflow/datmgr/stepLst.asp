<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i

set rsOra = createObject("adodb.recordset")

sqlOra = "select * from stepMst where cNo not in ('--', '00') and cNo > '0' order by cNo"
rsOra.open sqlOra, connOra
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Step list</title></head>
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
            
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Step list</font></b></td>
            <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            <td bgcolor="#0000FF"> 
                
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow" align="center"> 
            <td nowrap width="1%"><b><font color="#000000">No</font></b></td>
            <td nowrap width="1%"><b><font color="#000000">Step No</font></b></td>
            <td nowrap><b><font color="#000000">Name</font></b></td>
            <td nowrap><b><font color="#000000">Description</font></b></td>
            <td nowrap width="1%">&nbsp;</td>
            <td nowrap width="1%">&nbsp;</td>
          </tr>
          <%
					if not rsOra.EOF then
						i = 1
						
						do while not rsOra.EOF
							'if (isNumeric(rsOra("cNo")) and rsOra("cNo") > "0") or getAuthV2(session("sysauth"), AUTH_SYS_ADMIN) then
					%>
          <tr bgcolor="<%
					if i mod 2 = 0 then
						response.write "#FFCC00"
					else
						response.write "#FFFFCC"
					end if
					%>" valign="top"> 
            <td align="center" width="1%" nowrap><font color="#000000"><%=i%></font></td>
            <td align="center" width="1%" nowrap><font color="#000000"><%=rsOra("cNo")%></font></td>
            <td nowrap><font color="#000000"><%=showData(rsOra("vName"))%></font></td>
            <td width="100%"><font color="#000000"><%=showData(rsOra("vDesc"))%></font></td>
            <td align="center" nowrap width="1%"> 
              <%'if isNumeric(rsOra("cNo")) and rsOra("cNo") > "0" then%>
              <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editStepFrm.asp?stepNo=<%=rsOra("cNo")%>&sname=<%=server.urlEncode(rsOra("vName"))%>&sdesc=<%=server.urlEncode(rsOra("vDesc") & "")%>'">
              <%'end if%>
            </td>
            <td align="center" nowrap width="1%"> 
              <%'if isNumeric(rsOra("cNo")) and rsOra("cNo") > "0" then%>
              <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete step?\n\n\'<%=rsOra("cNo") & ": " & replace(replace(rsOra("vName"), "'", "\'"), QUOTE, "\" & QUOTE)%>\'')) {window.location.href = 'cgi-bin/delStep.asp?stepNo=<%=rsOra("cNo")%>'}">
              <%'end if%>
            </td>
          </tr>
          <%
								i = i + 1
							'end if
							
							rsOra.moveNext
						loop
					else
					%>
          <tr> 
            <td align="center" colspan="6" bgcolor="#FFFFFF"><br>
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
                <input type="button" value="Add step..." onClick="javascript: window.location.href = 'addStepFrm.asp'">
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