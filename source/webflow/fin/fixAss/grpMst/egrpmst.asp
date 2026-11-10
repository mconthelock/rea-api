<%Option Explicit%>

<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i

set rsOra = createObject("adodb.recordset")

sqlOra = "select " _
& "grpCode, grpdesc " _
& "from fxa_grpMst where grpCode ='" & request("grpCode") & "'"
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
<%
if rsOra.EOF then
	response.write "Data not found!"
else
%>
<form name="frmMain" method="get" action="upgrpmst.asp">
	<input type="hidden" name="grpCodeO" value="<%=rsOra("grpCode")%>">    
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
            <td align="right" valign="bottom"><img src="/form/imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/t.gif" width="100%" align="center"><img src="/form/imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="/form/imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr>
            <td background="/form/imgs/border/book/l.gif" align="right">&nbsp;</td>
            
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Edit Group Master</font></b></td>
            <td background="/form/imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td>
            
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr> 
            <td bgcolor="#FFCC00" nowrap width="0" valign="top"><b><font color="#000000">Group Code</font></b></td>
            <td bgcolor="#FFFFCC" nowrap width="100%"><input type="text" name="grpCode" value="<%=rsOra("grpCode")%>"></td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" nowrap width="0" valign="top"><b><font color="#000000">Group Description</font></b></td>
            <td bgcolor="#FFFFCC" nowrap width="100%"><font color="#000000"> 
              <textarea name="grpName" cols="60" rows="3"><%=rsOra("grpdesc")%></textarea>
              </font></td>
          </tr>
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
                
        <input type="submit" value="Save change" >
                
        <input type="button" value="Cancel" onClick="javascript: window.location.href = 'grpMst.asp'">
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
<%
end if

rsOra.close
%>
<!--#include virtual="/form/inc/copyRight.inc"-->
</body>
</html>

<%
set rsOra = nothing
%>