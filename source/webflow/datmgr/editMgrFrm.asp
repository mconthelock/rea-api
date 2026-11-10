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
<form name="frmMain" method="get" action="cgi-bin/editMgr.asp">
	<input type="hidden" name="empNo" value="<%=request.queryString("e")%>">
	<input type="hidden" name="orgNo" value="<%=request.queryString("o")%>">
	<input type="hidden" name="pos" value="<%=request.queryString("p")%>">
	<input type="hidden" name="co" value="<%if request.queryString("c") = "" then response.write "0" else response.write request.queryString("c") end if%>">
	<input type="hidden" name="headNo" value="<%=request.queryString("h")%>">
	<input type="hidden" name="orgNo1" value="<%=request.queryString("o1")%>">
	<input type="hidden" name="pos1" value="<%=request.queryString("p1")%>">
	<input type="hidden" name="co1" value="<%if request.queryString("c1") = "" then response.write "0" else response.write request.queryString("c1") end if%>">
	<input type="hidden" name="selDiv" value="<%=request.queryString("selDiv")%>">
	<input type="hidden" name="selDep" value="<%=request.queryString("selDep")%>">
	<input type="hidden" name="selSec" value="<%=request.queryString("selSec")%>">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr valign="bottom"> 
      <td align="right"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif">&nbsp;</td>
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2" dwcopytype="CopyTableRow">
          <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Manager 
              No :</font></b></td>
            <td bgcolor="#FFFFCC"> 
              <input type="text" name="txtMgrNo" value="<%=request.queryString("h")%>">
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Organize:</font></b></td>
            <td bgcolor="#FFFFCC"> 
              <%showOrgList "name==selOrg1&&misc==noBlankLine&&selected==" & request.queryString("o1")%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Position:</font></b></td>
            <td bgcolor="#FFFFCC"> 
              <%showPosList "name==selPos1&&misc==noBlankLine&&selected==" & request.queryString("p1")%>
              <input type="checkbox" name="chkCo1" value="1" <%if request.queryString("c1") <> "0" then response.write "checked" end if%>>
              Co-position </td>
          </tr>
        </table>
      </td>
      <td align="left" background="../imgs/border/book/r.gif">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif">&nbsp;</td>
      <td background="../imgs/border/book/bg.gif">&nbsp;</td>
      <td align="left" background="../imgs/border/book/r.gif">&nbsp;</td>
    </tr>
    <tr valign="top">
      <td align="right" background="../imgs/border/book/l.gif">&nbsp;</td>
      <td bgcolor="#0000FF">
		<table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Emp 
              No:</font></b></td>
            <td bgcolor="#FFFFCC"> 
              <input type="text" name="txtEmpNo" value="<%=request.queryString("e")%>">
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Organize:</font></b></td>
            <td bgcolor="#FFFFCC"> 
              <%showOrgList "name==selOrg&&misc==noBlankLine&&selected==" & request.queryString("o")%>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Position:</font></b></td>
            <td bgcolor="#FFFFCC"> 
              <%showPosList "name==selPos&&misc==noBlankLine&&selected==" & request.queryString("p")%>
              <input type="checkbox" name="chkCo" value="1" <%if request.queryString("c") <> "0" then response.write "checked" end if%>>
              Co-position </td>
          </tr>
        </table>
	  </td>
      <td align="left" background="../imgs/border/book/r.gif">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/bg.gif">&nbsp;</td>
      <td align="left" background="../imgs/border/book/r.gif"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/bg.gif" align="center"> 
        <input type="submit" value="Save">
        <input type="button" value="Back" onClick="javascript: window.location.href = 'mgrList.asp?selDiv=<%=request.queryString("selDiv") & "&selDep=" & request.queryString("selDep") & "&selSec=" & request.queryString("selSec")%>'">
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
<br>
<!--#include virtual="/sys/inc/copyRight.inc"-->
<br>
&nbsp; 
</body>
</html>
