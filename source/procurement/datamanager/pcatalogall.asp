<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	
    Dim  Rs,sSQL
	Dim sScriptPath,sCattype,Color,ExpType 
	Color="#F8D7D1"
	'suserid = Request("uid")
	'scorpcode = session("gbcorpcode")
	'Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	
	'sSQL = "SELECT  * FROM pCatalog where scatID = '"&Request.QueryString("sCatID")&"'  ORDER BY sCatID "
	sSQL = "SELECT * FROM PCATALOG A JOIN PCATTYPE B ON A.SCATTYPE = B.SCATTYPE "_
		&" WHERE SCATACC IS NOT NULL OR SEXPTYPE > '40' "_
		&" ORDER BY SEXPTYPE, A.SCATTYPE, SCATACC"
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 10
	
%>
<html>
<head>
<title>ExpensType</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
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
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Sub-Group in Expense Type 
                        </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
                <td width="15%" bgcolor="#FFD553" align="center" height="8"><A HREF="pdivisionedt.asp">
					<% if Mid(Session("sAuthor"),34,3) = "AED" Then %></A>
						<a href="pCatalogedt.asp"><img src="../images/iconadd2.gif" width="26" height="26" alt="ADD" border="0"></a>
					<% End if%>
				</td>
                <td width="9%" bgcolor="#FFD553" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="HELP"></td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="20%" bgcolor="#8B2718"><b><font color="#FFFFFF">Group in Expense Type</font></b></td>
                <td width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Expense ID</font></b></font></td>
				<td width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">SN</font></b></font></td>
                <td width="42%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Description</font></b></font></td>
                <td width="5%" bgcolor="#8B2718">&nbsp;</td>
                <td width="5%" bgcolor="#8B2718" align="right" valign="top"><img src="images/y05.jpg" width="12" height="12"></td>
              </tr>
              <% Do While Not RS.EOF %> 
              <tr bgcolor="<%=Color%>"> 
                <td bgcolor="<%=Color%>" height="25">
					<% If RS("sCATType") <> ExpType Then
						ExpType = RS("sCATType")
						Response.Write "<b>"&ExpType&"</b>- "&get_cattDesc(ExpType) 
					End IF %> 
				</td>
                <td bgcolor="<%=Color%>" height="25"><%=RS("sCatID")%>&nbsp;</td>
				<td bgcolor="<%=Color%>" height="25"><%=RS("sCatAcc")%>&nbsp;</td>
                <td bgcolor="<%=Color%>" height="25"><%=RS("sCatalog")%></td>
                <td bgcolor="<%=Color%>" align="center" height="25" valign="top"><A HREF="pdivisionedt.asp"><%
	if Mid(Session("sAuthor"),34,3) = "AED" Then
%></A><a href="pCatalogedt.asp?sCatID=<%=RS("sCatID")%>"><img src="../images/iconedit2.gif" width="26" height="26" border="0" alt="EDIT"></a><% End if%></td>
                <td width="9%" bgcolor="<%=Color%>" align="center" height="25"><%
	if Mid(Session("sAuthor"),34,3) = "AED" Then
%><a href="pCatalogdel.asp?sCatID=<%=RS("sCatID")%>"><img src="../images/icondelete2.gif" width="26" height="26" alt="DELETE" border="0"></a><% End if%> 
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
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="../datamanager/datamanager.asp">Data Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
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
<div align="right"></div>
</body>
</html>
