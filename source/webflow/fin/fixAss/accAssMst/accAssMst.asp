<%Option Explicit%>

<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i

set rsOra = createObject("adodb.recordset")

sqlOra = "select " _
& " accassCode, accassName, vOwner, sname " _
& "from fxa_accassMst, amec.aemployee a " _
& "where a.sempno(+)  = vOwner " _
& "order by accassCode"
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
            <td align="right" valign="bottom"><img src="/form/imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/t.gif" width="100%" align="center"><img src="/form/imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="/form/imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr>
            <td background="/form/imgs/border/book/l.gif" align="right">&nbsp;</td>
            
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">ACC ASSET MASTER</font></b></td>
            <td background="/form/imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td>
            
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow" align="center"> 
            <td nowrap width="1%"><b><font color="#000000">ASSET CODE</font></b></td>
            <td nowrap width="1%"><b><font color="#000000">ASSET NAME</font></b></td>
            <td nowrap><b><font color="#000000">CONTROLLER</font></b></td>
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
			 <td nowrap width="1%"><font color="#000000"><%=rsOra("accassCode")%></font></td>
            <td nowrap width="1%"><font color="#000000"><%=rsOra("accassName")%></font></td>
            <td nowrap width="100%"><font color="#000000">
			<%
				if not isnull(rsOra("vOwner")) then
					response.write "(" & rsOra("vOwner") & ") " & rsOra("sname")
				end if
			%>
			</font></td>			
            <td align="center" nowrap width="1%"> 
              <input type="button" value="Edit..." onClick="javascript: window.location.href = 'eccMst.asp?ccCode=<%=rsOra("accassCode")%>'">
            </td>
            <td align="center" nowrap width="1%"> 
              <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this topic?\n\n<%=rsOra("accassName")%>')) {window.location.href = 'dccMst.asp?ccCode=<%=rsOra("accassCode")%>'}">
            </td>
          </tr>
          <%
							i = i + 1
							rsOra.moveNext
						loop
					else
					%>
          <tr> 
            <td align="center" colspan="5" bgcolor="#FFFFFF"><br>
              <font class="alertMsg"><b>Data not found!</b></font><br>
              &nbsp;</td>
          </tr>
          <%
					end if
					
					rsOra.close
					%>
        </table>
            </td>
            <td background="/form/imgs/border/book/r.gif" align="left"><img src="/form/imgs/border/book/r.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/bg.gif">&nbsp;</td>
            <td background="/form/imgs/border/book/r.gif" align="left"><img src="/form/imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif"></td>
            <td align="center" background="/form/imgs/border/book/bg.gif"> 
        		<input type="button" value="Add..." onClick="javascript: window.location.href = 'accMst.asp'">
                <input type="button" value="Back" onClick="javascript: window.location.href = '../index.asp'">
            </td>
            <td background="/form/imgs/border/book/r.gif" align="left"><img src="/form/imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="/form/imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/b.gif"><img src="/form/imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="/form/imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
</form>
<!--#include virtual="/form/inc/copyRight.inc"-->
</body>
</html>

<%
set rsOra = nothing
%>