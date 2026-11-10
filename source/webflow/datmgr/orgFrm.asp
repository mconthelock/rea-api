<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/sys/inc/connUtil.inc"-->
<!--#include virtual="/sys/inc/userUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="frmMain" action="mgrList.asp" method="get">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr valign="bottom"> 
    <td align="right"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
    <td background="../imgs/border/book/t.gif" width="100%"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
    <td align="left"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
  </tr>
  <tr> 
    <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
    <td bgcolor="#0000FF"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr> 
          <td bgcolor="#FFCC00" width="1%"><b><font color="#000000">Division:</font></b></td>
          <td bgcolor="#FFFFCC"> 
            <%showDivList "name==selDiv&&selected==" & request.queryString("selDiv")%>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFCC00" width="1%"><b><font color="#000000">Department:</font></b></td>
          <td bgcolor="#FFFFCC"> 
            <%showDepList "name==selDep&&selected==" & request.queryString("selDep")%>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFCC00" width="1%"><b><font color="#000000">Section:</font></b></td>
          <td bgcolor="#FFFFCC"> 
            <%showSecList "name==selSec&&selected==" & request.queryString("selSec")%>
          </td>
        </tr>
      </table>
    </td>
    <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
  </tr>
  <tr valign="top">
    <td align="right" background="../imgs/border/book/l.gif"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
    <td background="../imgs/border/book/bg.gif">&nbsp;</td>
    <td align="left" background="../imgs/border/book/r.gif"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
  </tr>
  <tr valign="top"> 
    <td align="right" background="../imgs/border/book/l.gif"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
    <td background="../imgs/border/book/bg.gif" align="center"> 
        <input type="button" value="Add..." onClick="javascript: window.location.href = 'addMgrFrm.asp'">
        <input type="submit" value="Edit..." aonClick="javascript: window.location.href = 'mgrList.asp'">
      <input type="button" value="Back" onClick="javascript: window.location.href = 'index.asp'">
    </td>
    <td align="left" background="../imgs/border/book/r.gif"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
  </tr>
  <tr valign="top"> 
    <td align="right"><img src="../imgs/border/book/blc.gif" width="17" height="18"></td>
    <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" width="17" height="18"></td>
    <td align="left"><img src="../imgs/border/book/brc.gif" width="17" height="18"></td>
  </tr>
</table>
</form>
<br><!--#include virtual="/sys/inc/copyRight.inc"--><br>&nbsp;
</body>
</html>
