<% Option Explicit %>
<%
           ' Author : Ekkapan Kongchoo
%>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color,sexptype,keywords
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Color="#FFFFFF"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM pproduct  ORDER BY sprodcode "

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	TotalRows=Rs.RecordCount	
	RS.PageSize = 20
	PageSize=RS.PageSize
	TotalPage=Rs.PageCount
		if gPage="" then	
				PageNo=1
		else
				PageNo=gpage
		end if
	Select Case page
	Case "Prev"
			If gPage=1 then 
				PageNo=1
			else
				PageNo=gPage-1
			end if
	Case "Next"
			 If Cint(gPage)=Cint(TotalPage) then
					PageNo=TotalPage
			else
					PageNo=gPage+1
			end if
End Select
						
	
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
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr BGCOLOR="#3399CC"> 
                <td colspan="4" valign="top" height="8"><b><font color="#FFFFFF">Product 
                  </font></b></td>
              </tr>
              <tr bgcolor="#33CCFF"> 
                <td width="5%"><b><font color="#000000">No.</font></b></td>
                <td width="15%"><b><font color="#000000">Product Code</font></b></td>
                <td width="53%"><b><font color="#000000">Product Name</font></b></td>
                <td align="Right" width="10%"><b><font color="#000000">Price&nbsp;&nbsp;</font></b></td>
              </tr>
              <% '
RS.Absolutepage=PageNo
RowCount=1
Do While Not RS.EOF and RowCount<=PageSize
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="2"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="15%" bgcolor="<%=Color%>" height="2"><%=RS("sprodcode")%>&nbsp;</td>
                <td width="53%" bgcolor="<%=Color%>" height="2"><%=RS("seprodname")%></td>
                <td  align="right" width="10%" bgcolor="<%=Color%>" height="2"><%=RS("nPrice")%>&nbsp;&nbsp;</td>
              </tr>
              <%
if Color="#FFFFFF" then
		Color="#99CCFF"
else
		Color="#FFFFFF"
end if
RowCount=RowCount+1
 RS.MoveNext
Loop
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>

<%if TotalPage > 1 then%>
			<form action= "<%=sScriptPath%>" method="Post">
			<tr><td colspan="3" align="center" bgcolor="#FFFFFF">
			      <table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr><td align="center" bgcolor="#FFFFFF" width="2%"></td>
						<td align="left" bgcolor="#FFFFFF" width="45%">
<%if PageNo=1 then
			Response.write " "
else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev"><font color="#336699">Previous Page</font></a>
<%end if%>
</td>
<td align="right" bgcolor="#FFFFFF" width="45%">
<%
if PageNo=TotalPage Or PageNo=MaxPage then
	Response.write " "
else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next"><font color="#336699">Next Page</font></a>
<%End if%>
</td>
<td align="center" bgcolor="#FFFFFF" width="2%"></td></tr>
<tr>
<td align="center" bgcolor="#FFFFFF" colspan="5">
<%	
	For i=1 to TotalPage
		if Cint(PageNo)=Cint(i) then
%>
<font color="Black"><%=i%></font>
<%else%>
<A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxpage=<%=TotalPage%>">
<font color='#336699'><%=i%></font></a>
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
        <tr bgcolor="#FFFFFF"> 
          <td width="2%" align="left" valign="bottom" height="13"></td>
          <td width="96%" bgcolor="#FFFFFF" valign="middle" align="center" height="13"></td>
          <td width="2%" align="right" valign="bottom" height="13"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td> <br>
             <!--#include file="../inc/copyRight.inc"--></td>
        </tr>
      </table>
      </td>
  </tr>
</table>
<br>
<br>
<div align="right"></div>
</body>
</html>
