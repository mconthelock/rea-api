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

	sSQL = "SELECT  * FROM pPosition  ORDER BY sPoscode "

	'Response.Write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
%>
<html>
<head>
<title>Position</title>
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
          <td width="96%" height="2" bgcolor="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66"></font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr BGCOLOR="#3399CC"> 
                <td colspan="2" valign="top" height="8"><b><font color="#FFFFFF">Position</font></b> 
                </td>
              </tr>
              <tr bgcolor="#33CCFF"> 
                <td width="20%"><b><font color="#000000">Code</font></b></td>
                <td width="80%"><b><font color="#000000">Description</font></b></td>
              </tr>
              <% Do While Not RS.EOF 
			If RS("sPoscode")<> "00" then
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="20%" bgcolor="<%=Color%>" height="2"><%=RS("sposcode")%>&nbsp;</td>
                <td width="80%" bgcolor="<%=Color%>" height="2"><%=RS("sposition")%></td>
              </tr>
              <%
End IF
if Color="#FFFFFF" then
		Color="#99CCFF"
else
		Color="#FFFFFF"
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
          <td width="96%" bgcolor="#FFFFFF" valign="middle" align="center"><br>
</td>
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
