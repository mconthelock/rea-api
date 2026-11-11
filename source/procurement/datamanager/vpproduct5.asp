<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	'########## DERMPAN KHOOMRAKSA ##########
	'I change formatNumber(xxx, y) to formatNumber(xxx, 3) on 11-NOV-2002
	'########## END DERMPAN KHOOMRAKSA ##########
	
	' Author: Ekkapan Kongchoo
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color,sExptype,keywords,sCattype,sCatalog
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,sCatID
	Color="#F8D7D1"
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

if Request("status")="del" then
	sExptype=Session("ListProduct")
	scattype=Session("ListCattype")
	sCattype=Session("sCattype")
	MaxPage=Session("MaxPage")
	page=Session("page")
	gPage=Session("PageNo")
	keywords=Session("SearchProduct")
else
	sExptype=Request("ListProduct")
	scattype=Request("ListCattype")
	if Request("sCattype") = "" then
	sCattype=Request("ListCattype")
	else
		sCattype=Request("sCattype")
	end if
	MaxPage=Request("MaxPage")
	page=Request("page")
	gPage=Request("PageNo")
	keywords = TRIM( Request( "SearchProduct" ) )

end if

	'sSearchProduct=Request("SearchProduct")

if keywords="" then
	sSQL = "SELECT  * FROM pproduct Where sCattype='" & sCattype & "'" &_
		" and (sTprodname LIKE '%" &  keywords & "%' " &_
		" OR sEprodname Like '%" & keywords & "%' " &_
		" OR sprodcode Like '%" & keywords & "%' " &_
 		") ORDER BY sprodcode"
else
	sSQL = "SELECT  * FROM pproduct Where sTprodname LIKE '%" &  keywords & "%' " &_
		" OR sEprodname Like '%" & keywords & "%' " &_
		" OR sprodcode Like '%" & keywords & "%' " &_
 		" ORDER BY sprodcode"
end if
'	response.write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
	TotalRows=Rs.RecordCount	
	RS.PageSize = 30
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
			 If Cint(gPage)>=Cint(TotalPage) then
					PageNo=TotalPage
			else
					PageNo=gPage+1
			end if
End Select

%>

<html>
<head>
<title>Product :: <% Show_Pexptype(sExptype)%><%=Request("sCatdesc")%></title>
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../security/images/y02.jpg" width="12" height="12"></td>
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
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product 
                        :: <% Show_Pexptype(sExptype)%></font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
                <td width="10%" bgcolor="#FFD553" align="center" height="8">
		<%if Mid(Session("sAuthor"),10,3)="AED" then %>
				<a href="pproductedt.asp?sexptype_hid=<%=sexptype%>"><img src="../images/iconadd2.gif" width="26" height="26" alt="ADD" border="0"></a></td>
		<%End if%>
                <td width="15%" bgcolor="#FFD553" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="HELP"></td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="4%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product 
                  Code</font></b></font></td>
                <td width="43%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product 
                  Name</font></b></font></td>
                <td align="center" width="18%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">UPrice</font></b></font></td>
                <td width="10%" bgcolor="#8B2718">&nbsp;</td>
                <td width="15%" bgcolor="#8B2718" align="right" valign="top"><img src="images/y05.jpg" width="12" height="12"></td>
              </tr>
              <%If  RS.EOF then 
       		Response.Write  " <tr bgcolor='#F1B3A7'> " & vbCrLf
		       Response.Write  " <td colspan='3' valign='top'>ไม่มีรายการการที่ค้นหา"
	       	Response.Write  "</td></tr> "& vbCrLf
		else
				RS.Absolutepage=PageNo
				RowCount=1
				Do While Not RS.EOF and RowCount<=PageSize %> 
              <tr bgcolor="<%=Color%>"> 
                <td width="4%" bgcolor="<%=Color%>" height="2"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="10%" bgcolor="<%=Color%>" height="2"><%=RS("sprodcode")%>&nbsp;</td>
                <td width="43%" bgcolor="<%=Color%>" height="2"><a href ="vpproductdesc.asp?sprodcode=<%=RS("sProdcode")%>&sexptype=<%=RS("sexptype")%>"><%=RS("seprodname")%></a></td>
                <td  align="right" width="18%" bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 3)%></td>
                <td width="10%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
		<%if Mid(Session("sAuthor"),10,3)="AED" then %>
				<%response.write "<a href='pproductedt.asp?sprodcode="&Rs("sprodcode")&"&sexptype=" &RS("sexptype")&"&svendcode="&RS("svendcode")&"'>"%><img src="../images/iconedit2.gif" width="26" height="26" border="0" alt="EDIT"><%Response.Write("</a>")%></td>
		<%End if%>
                <td width="15%" bgcolor="<%=Color%>" align="center" height="2">
		<%if Mid(Session("sAuthor"),10,3)="AED" then %>
				<%response.write "<a href='pproductdel.asp?sprodcode="&Rs("sprodcode")&"&sexptype=" &RS("sexptype")&"&svendcode="&RS("svendcode")&"'>"%><img src="../images/icondelete2.gif" width="26" height="26" alt="DELETE" border="0"><%Response.Write("</a>")%> 
		<%End if%>
				<%Session("ListProduct")=sExptype%>
				<%Session("ListCattype")=scattype%>
				<%Session("sCatID")=sCatID%>
				<%Session("SearchProduct")=keywords%>
				<%Session("Pageno")=pageno%>
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
	end if
	Rs.close
	set Rs=nothing
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <%if TotalPage > 1 then%> 
        <form action= "<%=sScriptPath%>" method="Post">
          <tr> 
            <td colspan="4" align="center" bgcolor="#FFD553"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif" WIDTH="8" HEIGHT="8"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%
							if PageNo=1 then
											Response.write "  "
							else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListProduct=<%=sExptype%>&ListCattype=<%=sCattype%>&sCatID=<%=sCatID%>&SearchProduct=<%=keywords%>&page=Prev">Previous 
                    Page</a> <%					End if
%> </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <%
						If PageNo=TotalPage Or PageNo=MaxPage Then
							Response.write " "
						else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListProduct=<%=sExptype%>&ListCattype=<%=sCattype%>&sCatID=<%=sCatID%>&SearchProduct=<%=keywords%>&page=Next">Next 
                    Page</a> <%End if%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
						For i=1 to TotalPage
							If Cint(PageNo)=Cint(i) then
%> <Font color="red"><%=i%></Font> 
                    <%Else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&ListProduct=<%=sExptype%>&ListCattype=<%=sCattype%>&sCatID=<%=sCatID%>&SearchProduct=<%=keywords%>&MaxPage=<%=TotalPage%>"><%=i%></a> 
                    <%				
							End if	
							Next
 %> </td>
                </tr>
              </table>
              <%End if%> 
            </td>
          </tr>
        </form>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom" height="13"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center" height="13"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="../datamanager/datamanager.asp">Data Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
            </font></td>
          <td width="2%" align="right" valign="bottom" height="13"><img src="../security/images/y04.jpg" width="12" height="12"></td>
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
