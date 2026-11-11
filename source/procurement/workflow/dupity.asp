<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,sEmpapp
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
		  sSQL = "SELECT sEmpapp FROM  pappstep WHERE sEmpno='" & Session("sEmpno") & "'" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
IF Not RS.EOF then
		Do while Not Rs.EOF
			sEmpapp=Rs("sEmpapp")
		Rs.MoveNext
		Loop
Else
	Rs.Close
		  sSQL = "SELECT sEmpapp FROM  pexpstep WHERE sEmpno='" & Session("sEmpno") & "'" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Do while Not Rs.EOF
			sEmpapp=Rs("sEmpapp")
		Rs.MoveNext
		Loop
End if
		Rs.Close
		Set Rs=Nothing
%>
<html>
<head>
<title>Edit workflow Data</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2">&nbsp;</td>
    <td width="63%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="1" height="1"></td>
    <td width="36%" height="2" align="center"><img src="../images/spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td width="1%" height="2"><img src="../images/spacer.gif" width="1" height="1"></td>
    <td colspan="2" height="2" align="center" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="datamanager/../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#6666CC"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Deputize Approve </font></b></font></td>
                <td width="8%" bgcolor="#3838AB" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#E6E6F7" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="dup.asp">
                    <table width="75%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="27%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Deputy
                          Employee no</b>.</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                          </font></td>
                        <td width="73%"> 
                          <input type="text" name="sEmpno" size="12" maxlength="12" value="<%=sEmpapp%>">
                        </td>
                      </tr>
                    </table>
                    <br>
                    <input type="submit" name="Submit" value="Save">
                    <input type="Reset" name="Submit" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="workflow/workflow.htm">WorkFlow Manager</a></font></p>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td> <br>
            <!--#include file="../inc/copyRight.inc"--></td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
<br>
<br>
</body>
</html>
