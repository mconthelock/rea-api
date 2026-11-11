<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<!-- #include file ="Conn.asp" -->
<!-- #include file="fcommon.asp" -->
<!-- #include file="List.asp" -->

<%
	' Author: Ekkapan Kongchoo
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color,sexptype,keywords
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM pproduct  ORDER BY sprodcode "

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation

 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

%>

<html>
<head>
<title>Product </title>
<meta http-equiv="Content-Type" content="text/html; charset=">
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top"><b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td colspan="4" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">ORGANIZATION 
                        DATA :</font></b></font></b></font> <font color="#FFFFFF" size="2">Product 
                        </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
                <td width="8%" bgcolor="#FFD553" align="center" height="8"><%
	if Mid(Session("sAuthor"),10,1) = "A" Then
%><a href="pproductedt.asp"><img src="../images/iconadd2.gif" width="26" height="26" alt="ADD" border="0"></a><% End if%></td>
                <td width="9%" bgcolor="#FFD553" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="HELP"></td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product 
                  Code</font></b></font></td>
                <td width="40%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product 
                  Name</font></b></font></td>
                <td align="center" width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Expense Type</font></b></font></td>
                <td width="8%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Group</font></td>
                <td width="9%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Sub-Group</font></td>

              </tr>
              <% '
If Not Rs.EOF then

RowCount=1
Do While Not RS.EOF 
IF Instr(RS("sProdcode"),"""")<>"0" or Instr(RS("sEprodname"),"""")<>"0" then

%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="25"><%=RowCount%>&nbsp;</td>
                <td width="15%" bgcolor="<%=Color%>" height="25"><%=RS("sprodcode")%>&nbsp;</td>
                <td width="53%" bgcolor="<%=Color%>" hheight="25"><a href ="../datamanager/pproductdesc.asp?sprodcode=<%=RS("sProdcode")%>&sexptype=<%=RS("sexptype")%>"><%=RS("seprodname")%></a></td>
                <td  align="right" width="10%" bgcolor="<%=Color%>" height="25"><%=RS("sExptype")%></td>
                <td width="8%" bgcolor="<%=Color%>" align="center" height="25" valign="top"><%=RS("sCatType")%> </td>
                <td width="9%" bgcolor="<%=Color%>" align="center" height="25"> <%=RS("sCatID")%></td>
              </tr>
              <%
RowCount=RowCount+1
end if
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if

 RS.MoveNext
Loop
End if
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <%if TotalPage > 1 then%> 
        <form action= "<%=sScriptPath%>" method="Post">
          <tr> 
            <td colspan="3" align="center" bgcolor="#FFD553"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%if PageNo=1 then
			Response.write " "
else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev">Previous 
                    Page</a> <%end if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <%
if PageNo=TotalPage Or PageNo=MaxPage then
	Response.write " "
else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next">Next 
                    Page</a> <%End if%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="5"> <%	
	For i=1 to TotalPage
		if Cint(PageNo)=Cint(i) then
%> <font color="Red"><%=i%></font> 
                    <%else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxpage=<%=TotalPage%>"><%=i%></a> 
                    <%	
end if
Next 
%> 
                </tr>
              </table>
              <%End if%> 
            </td>
          </tr>
        </form>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom" height="13"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center" height="13"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()"><br>
            back</a> | <a href="../datamanager/datamanager.asp">Data Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
            </font></td>
          <td width="2%" align="right" valign="bottom" height="13"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td> <br>
           <!--#include file="inc/copyRight.inc"-->
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
