<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Dim sPrno,nItem,sExptype,sProdcode,keywords
	keywords=Ucase(Trim(Request("searchvender")))
	sProdcode=Request("sProdcode")
	sPrno = Request("sPrno")
	nItem = Request("nItem")
	sExptype = Request("sExptype")
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

if keywords="" then
	sSQL = "SELECT  * FROM pvender where sVendCode like '6%' and pvender.sename not like '(DO NOT USE)%' ORDER BY svendcode "
else
	sSQL = "SELECT  * FROM pvender where (sVendcode like '%" & keywords & "%'" &_
	" or sEname like '%" & keywords & "%'" &_ 
	" or sTname like '%" & keywords & "%'" &_
	" ) and pvender.sename not like '(DO NOT USE)%' ORDER BY svendcode "
end if
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
	'Response.write sSQL
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	TotalRows=Rs.RecordCount	
	RS.PageSize = 15
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
<head orig="undefined">
<title>Show Vendors </title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function CloseWindow_onClick() { //v2.0
  window.close();
}
//-->
</Script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
'			Response.write " sProdcode = "&sProdcode
'			Response.write " sVendcode = "&sVendcode
'			Response.write " sExptype = "&sExptype
'			Response.write "<BR>"
%>
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
                <td colspan="5" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendors 
                        </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="4%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="21%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendor 
                  Code</font></b></font></td>
                <td width="54%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendor 
                  Name</font></b></font></td>
                <td BGCOLOR="#993300" WIDTH="21%"> </td>
              </tr>
    <%If  RS.EOF then 
       		Response.Write  " <tr bgcolor='#F1B3A7'> " & vbCrLf
		       Response.Write  " <td colspan='4' valign='top'>ไม่มีรายการการที่ค้นหา"
	       	Response.Write  "</td></tr> "& vbCrLf
	else
		RS.Absolutepage=PageNo
		RowCount=1
		Do While Not RS.EOF and RowCount<=PageSize
%> 
              <tr bgcolor="<%=Color%>"> 
                <td bgcolor="<%=Color%>" height="10" width="4%"> 
                  <DIV Align="Right"><%=((PageNo-1)*RS.PageSize)+RowCount%>.&nbsp;</DIV></td>
                <td bgcolor="<%=Color%>" height="10" width="20%">&nbsp;&nbsp;<%=RS("sVendcode")%></td>
                <td bgcolor="<%=Color%>" height="10" width="55%">&nbsp;&nbsp; 
                  <a href='../Datamanager/pProductedt.asp?sVendcode=<%=RS("sVendcode")%>&sProdcode=<%=sProdcode%>&sExptype=<%=sExptype%>&Mode=Search' Target="num"><%=RS("sEname")%></a></td>
                <td WIDTH="21%" HEIGHT="10"> 
                  <DIV ALIGN="CENTER"><a href ='VenderDetail.asp?sVendcode=<%=RS("sVendcode")%>'>Show 
                    Detail</a></DIV>
                </td>
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
<%end if%>            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>

<%if TotalPage > 1 then%>
			<form action= "<%=sScriptPath%>" method="Post">
			<tr><td colspan="4" align="center" bgcolor="#FFD553">
			      <table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr><td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
						<td align="left" bgcolor="#FFD553" width="45%">
<%
			If PageNo=1 then
				response.write " "
			Else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&sProdcode=<%=sProdcode%>&sExptype=<%=sExptype%>&Mode=Search&searchvender=<%=keywords%>">Previous Page</a>
<%End if%>
</td>

<td align="right" bgcolor="#FFD553" width="45%">
<% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%>
<A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&sProdcode=<%=sProdcode%>&sExptype=<%=sExptype%>&Mode=Search&searchvender=<%=keywords%>">Next Page</a>
<%End if%>
</td>
<td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td></tr>
<tr>
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) Then
%> 
                    <Font color="Red"><%=i%></font> 
                    <%else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&sProdcode=<%=sProdcode%>&sExptype=<%=sExptype%>&Mode=Search&searchvender=<%=keywords%>"><%=i%></a> 
                    <%
	End if
	Next %></td>
</tr>
</table>
<%End if%>


			</td></tr>
			</form>

        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="Javascript:window.close()"> Close </a></font></td>
          <td width="2%" align="right" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>

      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
           <td> <br>
            <!--#include file="../inc/copyRight.inc"-->
		   </td>
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
