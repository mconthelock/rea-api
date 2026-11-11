<html>
<head>
<title>Home Procurement</title>
<link rel="stylesheet" href="styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="2%" height="2"><img src="images/ccb.gif" width="10" height="10"></td>
    <td width="60%" height="2" align="center" valign="top"><img src="images/spacer.gif" width="9" height="8"></td>
    <td width="38%" height="2" align="center"><img src="images/spacer.gif" width="9" height="8"></td>
  </tr>
  <tr> 
    <td width="2%" height="296"><img src="images/spacer.gif" width="10" height="10"></td>
    <td width="60%" height="296" align="center" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="97%" align="left" valign="top" height="30"><b><font face="Arial, Helvetica, sans-serif" size="+1">Procurement 
            Workflow System</font></b></td>
          <td width="3%" align="right">&nbsp;</td>
        </tr>
        <tr> 
          <td width="97%" align="left" valign="top" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66"><img src="images/y01.jpg" width="12" height="12"></font></b></font></td>
          <td width="3%" align="right">&nbsp;</td>
        </tr>
        <tr> 
          <td width="97%" align="center" valign="middle" bgcolor="#FFD553" height="270"> 
            <img src="images/okpic001.gif" width="300" height="250"><br>
          </td>
          <td width="3%" align="right">&nbsp;</td>
        </tr>
        <tr> 
          <td width="97%" align="right" valign="bottom" bgcolor="#FFD553"><img src="images/y04.jpg" width="12" height="12"></td>
          <td width="3%" align="right">&nbsp;</td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
    <td width="38%" align="center" valign="top" height="296"> <br>
      <br>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td width="96%" height="3" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="3"><img src="images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" height="100">&nbsp;</td>
          <td width="96%" align="left" valign="top" height="100"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr> 
                <td bgcolor="#28AA56"><font face="Verdana, Arial, Helvetica, sans-serif" color="#EEEEEE"><b>Today 
                  is...<br>
                  <%=FormatDateTime(Now(),1)%>&nbsp;<%=FormatDateTime(Now(),4)%>&nbsp;</b></font></td>
              </tr>
              <tr> 
                <td><br>
                  <b><font color="#FF0000">Welcome ....</font></b> <br>
                  <b><%=Session("sName")%> &nbsp;</b></td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
              </tr>
              <tr> 
                <td bgcolor="#28AA56"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Workflow 
                  Task </font></b></font></td>
              </tr>
              <tr> 
                <td>
<table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr> 
                      <td ><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"><img src="images/plus.gif" width="9" height="9"></font> 
                        <font face="Verdana, Arial, Helvetica, sans-serif"><a href="security/security.asp" target="num" ><b>Security 
                        Manager</b></a></font></td>
                    </tr>
                    <tr> 
                      <td height="2" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"><img src="images/plus.gif" width="9" height="9"></font> 
                        <font face="Verdana, Arial, Helvetica, sans-serif"><b><a href="datamanager/datamanager.asp" target="num" >Data 
                        Manager</a></b></font></td>
                    </tr>
                    <tr> 
                      <td height="2" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"><img src="images/plus.gif" width="9" height="9"></font> 
                        <font face="Verdana, Arial, Helvetica, sans-serif"><b><a href="workflow/workflow.asp" target="num" >Workflow 
                        Manager</a></b></font></td>
                    </tr>
                    <tr> 
                      <td height="2" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#DB4B37"><img src="images/plus.gif" width="9" height="9"></font> 
                        <font face="Verdana, Arial, Helvetica, sans-serif"><b><a href="report/report.asp" target="num" >Report</a></b></font></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <font color="#006699"> </font></td>
          <td width="2%" height="100">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td width="1%"> <img src="images/spacer.gif" width="10" height="10"><br>
    </td>
    <td width="99%"> <br>
      <!--#include file="inc/copyRight.inc"-->
	  </td>
  </tr>
</table>
<br>
</body>
</html>
