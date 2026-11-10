<%
Option Explicit
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="frmMain" method="get" action="cgi-bin/addStep.asp">
    
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr>
      <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Add step</font></b></td>
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
                  <td bgcolor="#FFFFCC"> <font color="#000000"> 
                    <input type="text" name="txtName" onKeyUp="javascript: chkLng(document.frmMain.txtName, 30, 'Remark is no longer than 30 charactors')">
                    </font></td>
                </tr>
                <tr> 
                  <td width="1%" nowrap valign="top" bgcolor="#FFCC00"><b><font color="#000000">Description:</font></b></td>
                  <td valign="top" bgcolor="#FFFFCC"> 
                    <textarea name="txtDesc" cols="50" rows="5" onKeyUp="javascript: chkLng(document.frmMain.txtDesc, 256, 'Remark is no longer than 256 charactors')"></textarea>
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
        <input type="submit" value="Add">
        <input type="button" value="Back" onClick="javascript: window.location.href='stepLst.asp'">
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
	document.frmMain.txtName.focus();
</script>
</body>
</html>
