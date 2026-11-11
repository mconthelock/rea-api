<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->

<%
	' Author: Ekkapan Kongchoo
	
    Dim  Rs,sSQL
	Dim sScriptPath,sexptype,Color,exptype

  Color="#F8D7D1"
	'suserid = Request("uid")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  * FROM vpexpstep  ORDER BY sexptype,nstep "

	'Response.Write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
	
%>
<html>
<head>
<title>Expense Type Approval Step</title>
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</Script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2">&nbsp;</td>
    <td width="63%" height="2" align="center" valign="top"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td width="36%" height="2" align="center"><img src="../security/images/spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td width="1%" height="2"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td colspan="2" height="2" align="center" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><FONT COLOR="#FFCC66">..</FONT></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../security/images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"><b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td colspan="3" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="97%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> 
                        Approval Method : Expense Type Approval Step </font></b></font></td>
                      <td width="3%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
                <td width="45" bgcolor="#FFD553" align="center" height="8"><A HREF="pdivisionedt.asp"><%
	if Mid(Session("sAuthor"),4,1) = "A" Then
%></A><a href="../datamanager/pexpstepedt.asp"><img src="../images/iconadd2.gif" width="26" height="26" alt="ADD" border="0"><% End if%></a></td>
                <td width="26" bgcolor="#FFD553" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="HELP"></td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td bgcolor="#8B2718" width="163"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Expense 
                  Type</font></b></font></td>
                <td bgcolor="#8B2718" width="57"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Step</font></b></font></td>
                <td bgcolor="#8B2718" width="224"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Description</font></b></font></td>
                <TD WIDTH="45" BGCOLOR="#8B2718">&nbsp;</TD>
                <TD WIDTH="26" BGCOLOR="#8B2718" VALIGN="top"> 
                  <DIV ALIGN="RIGHT"><IMG SRC="../images/y05.jpg" WIDTH="11" HEIGHT="12"></DIV>
                </TD>
              </tr>
              <% Do While Not RS.EOF %> 
              <tr bgcolor="<%=Color%>"> 
                <td bgcolor="<%=Color%>" height="25" width="163"> <%
if RS("sexptype")  <> exptype Then
	exptype = RS("sexptype")
   Response.write "<b>"&exptype& "</b> - "&RS("sexpdesc")
end if
%></td>
                <td bgcolor="<%=Color%>" height="25" width="57"><%=RS("nstep")%>&nbsp;</td>
                <td bgcolor="<%=Color%>" height="25" width="224">
	<% 
	   if RS("sempno") <> "" Then
		 Response.Write "<b><A HREF=JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno="&RS("sempno")&"','ViewUser','width=500','height=250')>"& Rs("sEmpno") & "</a></b> - "
 	   end if
	%>

<%=RS("sdesc")%>&nbsp;</td>
                <td width="45" bgcolor="<%=Color%>" align="center" height="25" valign="top"><A HREF="pdivisionedt.asp"><%
	if Mid(Session("sAuthor"),5,1) = "E" Then
%></A><a href="pexpstepedt.asp?nstep=<%=RS("nstep")%>&sexptype=<%=RS("sexptype")%>"><img src="../images/iconedit2.gif" width="26" height="26" border="0" alt="EDIT"><% End if%></a></td>
                <td width="26" bgcolor="<%=Color%>" align="center" height="25"><A HREF="pdivisionedt.asp"><%
	if Mid(Session("sAuthor"),6,1) = "D" Then
%></A><a href="pexpstepdel.asp?nstep=<%=RS("nstep")%>&sexptype=<%=RS("sexptype")%>"><img src="../images/icondelete2.gif" width="26" height="26" alt="DELETE" border="0"><% End if%></a> 
                </td>
              </tr>
              <%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
 RS.MoveNext
Loop
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
            <br>
            </font></td>
          <td width="2%" align="right" valign="bottom"><img src="../security/images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td> <br>
            <!--#include file="../inc/copyRight.inc"-->
		   </td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
<br>
<br>
<div align="right">
  <hr noshade size="1">
</div>
</body>
</html>
