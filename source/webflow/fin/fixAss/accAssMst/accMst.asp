<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Form list</title></head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">
<form name="frmMain" method="get" action="accMsti.asp">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
            <td align="right" valign="bottom"><img src="/form/imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="/form/imgs/border/book/t.gif" width="100%" align="center"><img src="/form/imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="/form/imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr>
            <td background="/form/imgs/border/book/l.gif" align="right">&nbsp;</td>
            
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">AACC ASSET MASTER</font></b></td>
            <td background="/form/imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td>
            
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow"> 
            <td nowrap valign="top" width="0" bgcolor="#FFCC00"> <b><font color="#000000">ASSET CODE</font></b></td>
            <td nowrap width="100%" bgcolor="#FFFFCC"><b> 
              <input type="text" name="ccCode" maxlength="10">
              </b></td>
          </tr>
          <tr> 
            <td nowrap width="0" bgcolor="#FFCC00" valign="top"><b><font color="#000000">ASSET NAME</font></b></td>
            <td nowrap width="100%" bgcolor="#FFFFCC"> 
              <textarea name="ccName" cols="60" rows="5"></textarea>
            </td>
          </tr>
          <tr> 
            <td nowrap width="0" bgcolor="#FFCC00" valign="top"><b><font color="#000000">CONTROLLER</font></b></td>
            <td nowrap width="100%" bgcolor="#FFFFCC"> <input type="text" name="txtvowner"></td>
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
                
        <input type="submit" value="Add">
                
        <input type="button" value="Cancel" onClick="javascript: window.location.href = 'ccMst.asp'">
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