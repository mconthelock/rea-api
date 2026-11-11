<% Option Explicit %>
<%
           ' Author : Ekkapan Kongchoo
%>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
    Dim  Rs,sSQL
	Dim sScriptPath,sexptype,Color
		Color="#99CCFF"
	'suserid = Request("uid")
	'scorpcode = session("gbcorpcode")
	'Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  * FROM pdivision  ORDER BY sdivcode "

	'Response.Write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
%>
<html>
<head>
<title>Division</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2"></td>
    <td width="63%" height="2" align="center" valign="top"></td>
    <td width="36%" height="2" align="center"></td>
  </tr>
  <tr> 
    <td width="1%" height="2"></td>
    <td colspan="2" height="2" align="center" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFFFFF"> 
          <td align="left" valign="top" width="2%" height="2">&nbsp;</td>
          <td width="96%" height="2" bgcolor="#FFFFFF"><b><font color="#FFFFFF">..</font></b></td>
          <td align="right" valign="top" width="2%" height="2"></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr BGCOLOR="#3399CC"> 
                <td colspan="2" valign="top" height="8"><b><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#FFFFFF">Division</FONT></b> 
                </td>
              </tr>
              <tr bgcolor="#33CCFF"> 
                <td BGCOLOR="#33CCFF" width="20%"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#000000" size="2">Code</font></b></font></td>
                <td width="80%"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#000000" size="2">Description</font></b></font></td>
              </tr>
              <% Do While Not RS.EOF 
IF RS("sDivcode")<>"00" then
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="20%" bgcolor="<%=Color%>"><%=RS("sdivcode")%>&nbsp;</td>
                <td width="80%" bgcolor="<%=Color%>"><%=RS("sdivision")%></td>
              </tr>
              <%
End if
if Color="#99CCFF" then
		Color="#FFFFFF"
else
		Color="#99CCFF"
end if
 RS.MoveNext
Loop
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="2%" align="left" valign="bottom"></td>
          <td width="96%" bgcolor="#FFFFFF"><BR></td>
          <td width="2%" align="right" valign="bottom"></td>
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
<div align="right"></div>
</body>
</html>
