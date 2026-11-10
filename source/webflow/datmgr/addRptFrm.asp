<%
Option Explicit
%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->
<!--#include file="../inc/security.inc"-->
<!--#include file="../inc/userUtil.inc"-->
<!--#include file="../inc/formUtil.inc"-->

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
<form name="frmMain" method="get" action="cgi-bin/addRpt.asp">
  <input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Add report</font></b></td>
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
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Organize:</font></b></td>
                  <td bgcolor="#FFFFCC"><%showOrgList "name==selOrg"%></td>
                </tr>
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Name:</font></b></td>
                  <td bgcolor="#FFFFCC"> 
                    <input type="text" name="txtName" maxlength="120">
                  </td>
                </tr>
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Desc:</font></b></td>
                  <td bgcolor="#FFFFCC"> 
                    <input type="text" name="txtDesc" maxlength="256">
                  </td>
                </tr>
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">URL:</font></b></td>
                  <td bgcolor="#FFFFCC"> 
                    <input type="text" name="txtURL" maxlength="256">
                  </td>
                </tr>
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Show:</font></b></td>
                  <td bgcolor="#FFFFCC">
				  	<table cellpadding="0" cellspaciing="0" border="0">
						<tr>
							<td><input type="radio" name="optShow" value="1" checked></td>
		                    <td><font color="#000000">yes</font></td>
							<td><input type="radio" name="optShow" value="2"></td>
		                    <td><font color="#000000">no</font></td>
						</tr>
					</table>
					</td>
                </tr>
				<!--
                <tr> 
                  <td width="1%" nowrap bgcolor="#FFCC00"><font color="#000000"><b>Icon:</b></font></td>
                  <td bgcolor="#FFFFCC"> 
                    <input type="file" name="file">
                    <font color="#0000FF">(left blank to use default icon)</font> 
                  </td>
                </tr>
				-->
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
        <%'if mode = MODE_EDIT then%>
        <input type="submit" value="Add">
        <%'end if%>
        <input type="button" value="Back" onClick="javascript: window.location.href='rptLst.asp'">
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
