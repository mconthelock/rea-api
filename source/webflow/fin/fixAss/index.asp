<!--#include virtual="/sys/inc/as400.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->

<%
if request("no") <> "" and request("orgno") <> "" and request("y") <> "" then
	session("no") = request("no")
	session("orgNo") = request("orgNo")
	session("y") = request("y")
end if
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
if not getAuth(session("no"), session("orgNo"), session("y"), session("uid"), AUTH_FORM_ADMIN) then
	response.write "<font color=""#FF0000""><b>You don't have permission to access this page!</b></font>"
	response.end
end if
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
  <tr> 
    <td align="right" valign="bottom"><img src="/form/imgs/border/book/tlc.gif" width="17" height="18"></td>
    <td background="/form/imgs/border/book/t.gif" width="100%" align="center"><img src="/form/imgs/border/book/t.gif" width="17" height="18"></td>
    <td align="left" valign="bottom"><img src="/form/imgs/border/book/trc.gif" width="17" height="18"></td>
  </tr>
  <tr> 
    <td background="/form/imgs/border/book/l.gif" align="right">&nbsp;</td>
    <td bgcolor="#AACCEE" align="left"><b><font color="#000000">Fixed Asset</font></b></td>
    <td background="/form/imgs/border/book/r.gif" align="left">&nbsp;</td>
  </tr>
  <tr> 
    <td background="/form/imgs/border/book/l.gif" align="right"><img src="/form/imgs/border/book/l.gif" width="17" height="18"></td>
    <td bgcolor="#0000FF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"><a href="https://amecweb.mitsubishielevatorasia.co.th/form/finform/FIN-PCK/form/locmst">Location Master</a></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"><font color="#000000"><a href="accAssMst/accAssMst.asp">ACC ASSET MASTER</a></font></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"><font color="#000000"><a href="grpMst/grpMst.asp">Group Master</a></font></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"><font color="#000000"><a href="faMst/faMst.asp">Fixed Asset Master</a></font></td>
        </tr>
        <!--tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"> 
            <p><font color="#000000"><a href="seemf.asp">CONFIRM NEW FIXED ASSET</a></font></p>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"> 
            <p><font color="#000000"><a href="rtoMstf.asp">DISPOSAL ASSETS REQUEST FORM</a></font></p>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"> 
            <p><font color="#000000"><a href="rtoMstf.asp">ASSETS TRANSFERRED SHEET</a></font></p>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"> 
            <p><font color="#000000"><a href="rtoMstf.asp">ASSETS TRANSFERRED SHEET</a></font></p>
          </td>
        </tr -->					
        <tr> 
          <td bgcolor="#FFFFCC"> 
            <div align="center"><img src="../../imgs/r_arrow.gif" width="10" height="10"></div>
          </td>
          <td bgcolor="#FFFFCC"> 
            <p><font color="#000000"><a href="report/finFXSearch.asp">FIXED ASSETS REGISTER	</a></font></p>
          </td>
        </tr>
      </table>
    </td>
    <td background="/form/imgs/border/book/r.gif" align="left"><img src="/form/imgs/border/book/r.gif" width="17" height="18"></td>
  </tr>
  <tr> 
    <td align="right" valign="top"><img src="/form/imgs/border/book/blc.gif" width="17" height="18"></td>
    <td background="/form/imgs/border/book/b.gif"><img src="/form/imgs/border/book/b.gif" width="17" height="18"></td>
    <td align="left" valign="bottom"><img src="/form/imgs/border/book/brc.gif" width="17" height="18"></td>
  </tr>
</table>
</body>
</html>
