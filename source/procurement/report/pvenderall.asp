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
	Dim sScriptPath,svendcode,Color
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Color="#FFFFFF"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM pvender  ORDER BY svendcode "
	sSQL = "SELECT  * FROM pvender, ptermcode where pvender.stermcode = ptermcode.stermcode(+) ORDER BY svendcode"
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
<title>Vendors </title>
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
          <td width="96%" height="2" bgcolor="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66"></font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFFFFF"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr BGCOLOR="#3399CC"> 
                <td colspan="10" valign="top" height="8"><b><font color="#FFFFFF">Vendors 
                  </font></b> </td>
              </tr>
              <tr bgcolor="#FFFFFF"> 
                <td width="5%" bgcolor="#33CCFF"><B>No.</B></td>
                <td width="14%" bgcolor="#33CCFF"><b> Code</b></td>
                <td width="25%" bgcolor="#33CCFF"><b>Vendor Name</b></td>
                <td width="25%" bgcolor="#33CCFF"><b>Term Code</b></td>
                <td width="25%" bgcolor="#33CCFF"><b>City</b></td>
                <td width="25%" bgcolor="#33CCFF"><b>Contact</b></td>
                <td width="25%" bgcolor="#33CCFF"><b>Currency</b></td>
                <td width="32%" bgcolor="#33CCFF"><b>Address</b></td>
                <td width="13%" bgcolor="#33CCFF"><b>Tel.</b></td>
                <td width="11%" bgcolor="#33CCFF"><b>Fax.</b></td>
              </tr>
              <%
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize
%>
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="2">&nbsp;<%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="14%" bgcolor="<%=Color%>" height="2"><%=RS("sVendcode")%>&nbsp;</td>
                <td width="25%" bgcolor="<%=Color%>" height="2"><%=RS("sename")%></td>
                <td width="25%" bgcolor="<%=Color%>" height="2"><%=RS("sTermDesc")%></td>
                <td width="25%" bgcolor="<%=Color%>" height="2"><%=RS("sCity")%></td>
                <td width="25%" bgcolor="<%=Color%>" height="2"><%=RS("sContact")%></td>
                <td width="25%" bgcolor="<%=Color%>" height="2"><%=RS("sStdCur")%></td>
                <td width="32%" bgcolor="<%=Color%>" height="2"><%=RS("sEaddress")%>&nbsp;</td>
                <td width="13%" bgcolor="<%=Color%>" height="2"><%=RS("sTelno")%>&nbsp;</td>
                <td width="11%" bgcolor="<%=Color%>" height="2"><%=RS("sFax")%>&nbsp;</td>
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
			<tr><td colspan="4" align="center" bgcolor="#FFFFFF">
			      <table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr><td align="center" bgcolor="#FFFFFF" width="2%"></td>
						<td align="left" bgcolor="#FFFFFF" width="45%">
<%
			If PageNo=1 then
				response.write " "
			Else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev"><font color="#336699">Previous Page</font></a>
<%End if%>
</td>

<td align="right" bgcolor="#FFFFFF" width="45%">
<% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next"><font color="#336699">Next Page</font></a>
<%End if%>
</td>
<td align="center" bgcolor="#FFFFFF" width="2%"></td></tr>
<tr>
<td align="center" bgcolor="#FFFFFF" colspan="4">
<%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) Then
%>
<Font color="#000000"><%=i%></font>
<%else%>
			<A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>">
			<font color='#336699'><%=i%></font></a>
<%
	End if
	Next %></td>
</tr>
</table>
<%End if%>
<%
if rowCount > 1 then
	response.write "<br><table align=""center""><tr><td colspan=""2"" align=""center"">Download MS Excel file</td></tr><tr><td><input type=""button"" value=""Current page"" onClick=""javascript: window.location.href = 'pvenderalltoexcel.asp?s=" & server.urlEncode(sSql) & "&r=" & (pageNo * pageSize) - pageSize & "&r2=" & rowCount & "'"">&nbsp<input type=""button"" value=""All records"" onClick=""javascript: window.location.href = 'pvenderalltoexcel.asp?s=" & server.urlEncode(sSql) & "&r=0&r2=0'""></td></tr></table>"
end if
%>

			</td></tr>
			</form>
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
