<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo

   Dim  Rs,sSQL
	Dim sScriptPath,sVendcode,Color,keywords,sProdcode
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Dim sPrno,nItem,sExptype
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	sPrno = Request("sPrno")
	sProdcode = Request("sProdcode")
	sVendcode = Request("sVendcode")
	nItem = Request("nItem")
	sExptype= Request("sExptype")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM pProduct  ORDER BY sprodcode "
    'response.write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	TotalRows=Rs.RecordCount	
	RS.PageSize = 10
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
<title>Show Product </title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="3" topmargin="3" marginwidth="3" marginheight="3">
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
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td colspan="5" bgcolor="#DB4B37" valign="top" height="17"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td height="10"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product 
                        </font></b></font></td>
                      <td align="right" valign="top" height="10"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="4%" bgcolor="#8B2718">
                  <DIV ALIGN="CENTER"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></DIV>
                </td>
                <td width="28%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">&nbsp;&nbsp;Product 
                  Code</font></b></font></td>
                <td width="43%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">&nbsp;&nbsp;Product 
                  Name</font></b></font></td>
                <td align="center" width="13%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Price</font></b></font></td>
                <td width="12%" bgcolor="#8B2718" align="right" valign="top">&nbsp;</td>
              </tr>
              <% '
RS.Absolutepage=PageNo
RowCount=1
Do While Not RS.EOF and RowCount<=PageSize
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="4%" bgcolor="<%=Color%>">
					<DIV ALIGN="RIGHT">
						<%=((PageNo-1)*RS.PageSize)+RowCount%>.&nbsp;
					</DIV>
                </td>
                <td width="28%" bgcolor="<%=Color%>" height="10">&nbsp;&nbsp;<%=RS("sprodcode")%>&nbsp;</td>
                <td width="43%" bgcolor="<%=Color%>" hheight="25" HEIGHT="10">&nbsp;&nbsp;<a href=../workflow/ppritemedt.asp?nItem=<%=nItem%>&sExptype=<%=Rs("sExptype")%>&sPrno=<%=sPrno%>&sProdcode=<%=RS("sProdcode")%> Target="num"><%=RS("sEprodname")%></a> </td>
                <td  align="right" width="13%" bgcolor="<%=Color%>" height="10"><%=RS("nPrice")%></td>
                <td width="12%" bgcolor="<%=Color%>" align="center" height="10"><A HREF="ProductDetail.asp?sprodcode=<%=RS("sProdcode")%>&sexptype=<%=RS("sExptype")%>">Detail</A></td>
              </tr>
              <%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
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
			<tr><td colspan="3" align="center" bgcolor="#FFD553">
			      <table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr><td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
						<td align="left" bgcolor="#FFD553" width="45%">
<%if PageNo=1 then
			Response.write " "
else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&nItem=<%=nItem%>&sExptype=<%=sExptype%>&sPrno=<%=sPrno%>">Previous Page</a>
<%end if%>
</td>

<td align="right" bgcolor="#FFD553" width="45%">
<%
if PageNo=TotalPage Or PageNo=MaxPage then
	Response.write " "
else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&nItem=<%=nItem%>&sExptype=<%=sExptype%>&sPrno=<%=sPrno%>">Next Page</a>
<%End if%>
</td>
<td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td></tr>
<tr>
<td align="center" bgcolor="#FFD553" colspan="5">
<%	
	For i=1 to TotalPage
		if Cint(PageNo)=Cint(i) then
%>
<font color="Red"><%=i%></font>
<%else%>
			<A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxpage=<%=TotalPage%>&nItem=<%=nItem%>&sExptype=<%=sExptype%>&sPrno=<%=sPrno%>"><%=i%></a>
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
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center" height="13">
				<BR>
					<FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2">
						<A HREF="javascript:history.back()">back</A> 
			            | <A HREF="Javascript:window.close()"> Close </A>
					</FONT>
				<BR>
			</td>
          <td width="2%" align="right" valign="bottom" height="13"><img src="../images/y04.jpg" width="12" height="12"></td>
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
