<%
Option Explicit
%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->
<!--#include file="../inc/security.inc"-->

<%
Dim rsOra, sqlOra
Dim mode

mode = MODE_VIEW
set rsOra = createObject("adodb.recordset")
%>
<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
<title>Add organization</title>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<form name="frmMain" method="get" action="cgi-bin/addFrm.asp">
  <input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Add form</font></b></td>
      <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#FFFFFF"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#0000FF"> 
              <table width="100%" border="0" cellspacing="1" cellpadding="2">
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Name:</font></b></td>
                  <td bgcolor="#FFFFCC"> 
                    <input type="text" name="txtName" maxlength="60">
                  </td>
                </tr>
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Desc:</font></b></td>
                  <td bgcolor="#FFFFCC"> 
                    <input type="text" name="txtDesc" maxlength="256">
                  </td>
                </tr>
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Organization:</font></b></td>
                  <td bgcolor="#FFFFCC"> 
                    <%
							sqlOra = "select * from vOrgMst where vOrgNo <> '00' and vOrgNo in (select distinct vOrgNo from orgTree) order by cType desc, vOrgNo"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
								mode = MODE_EDIT
							%>
                    <select name="selOrg">
                      <%
								do while not rsOra.EOF
									response.write "<option value=" & QUOTE & rsOra("vOrgNo") & QUOTE & ">" & rsOra("vOrgNo") & ": " & rsOra("vName") & " (" & rsOra("vDesc") & ")</option>"
									rsOra.moveNext
								loop
							%>
                    </select>
                    <%
							else
							%>
                    <font color="#000000">Organization not found!</font> 
                    <%
							end if
							
							rsOra.close
							%>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
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
        <%if mode = MODE_EDIT then%>
        <input type="submit" value="Add">
        <%end if%>
        <input type="button" value="Back" onClick="javascript: window.location.href='frmLst.asp'">
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
<script language="JavaScript">
	document.frmMain.selOrg.focus();
</script>
</body>
</html>
