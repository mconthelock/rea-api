<!--#include file="inc/login.inc"-->

<html>
<head>
<link href="/sys/css/common.css" rel="stylesheet" type="text/css">
</head>
<body>
<center>
    <p><font color="#0000FF"><b> <br>
        </b></font></p>
    <p>&nbsp;</p>
    </center>
<form name="frmMain" action="cgi-bin/login.asp" method="post">
    <table border="0" cellspacing="1" cellpadding="2" align="center">
        <tr align="center"> 
            <td colspan="2"><img src="imgs/mitsu.gif" width="105" height="95"></td>
        </tr>
        <tr> 
            <td colspan="2">
                <center>
                    <font color="#0000FF"><b> Mitsubishi Elevator Asia Co.,Ltd.<br>
                    Information System department Electronics forms</b></font> 
                </center>
            </td>
        </tr>
    </table>
    <br>
    <table border="0" cellspacing="1" cellpadding="2"<%test%> align="center">
        <tr> 
            <td><font color="#000000"><b>User ID:</b></font></td>
            <td align="center"> 
                <input type="text" name="txtUid" size="10">
            </td>
            <td align="center">&nbsp;</td>
        </tr>
        <tr> 
            <td><b><font color="#000000">Password:</font></b></td>
            <td align="center"> 
                <input type="password" name="txtPwd" size="10">
            </td>
            <td align="center">
                <input type="submit" value="Login">
            </td>
        </tr>
    </table>
</form>
<script language="JavaScript">
document.frmMain.txtUid.focus();
</script>
</body>
</html>
