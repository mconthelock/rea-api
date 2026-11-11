<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo
	
   Dim  Rs,RS2,sSQL,sSQL2
	Dim sScriptPath,svendcode,Color,sExptype,keywords,vendcode
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
		Color="#F8D7D1"
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		Set Rs = Server.CreateObject("ADODB.RecordSet")
if request("status")="del" then
		sExptype=Session("ListVender")
		keywords =Ucase(Session("SearchVender"))
		gPage=Session("PageNo")
		MaxPage=Session("MaxPage")
else       	
		sExptype=Request("ListVender")
		keywords = Ucase(TRIM( Request( "SearchVender" ) ))
		page=Request("page")
		gPage=Request("PageNo")
		MaxPage=Request("MaxPage")
end if
'		sSQL = "SELECT distinct  svendcode  FROM  pProduct  where sexptype='" & sExptype & "'"
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
if  keywords=""  then
	sSQL = "SELECT  * FROM pvender  ORDER BY svendcode "

else 
	sSQL = "Select * from pVender where stname LIKE '%" &  keywords & "%' " &_
		" OR sename Like '%" & keywords & "%' " &_
		" OR sVendcode Like '%" & keywords & "%' " &_
		" Or sSearchkey Like '%" & keywords & "%' " &_
		" Order By sVendcode"
'	sSQL = "Select * from pVender where sVendcode Like '%" & keywords & "%' "
End if
	 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 30
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
<title>Vender :: <% Show_Pexptype(sExptype)%></title>
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
          <td width="96%" align="left" valign="top"><b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td colspan="5" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendor 
                        : <% Show_Pexptype(sExptype)%></font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
                <td width="8%" bgcolor="#FFD553" align="center" height="8">
		<%if Mid(Session("sAuthor"),7,3)="AED" then %>
				<a href="pvenderedt.asp?sexptype_hid=<%=sexptype%>"><img src="../images/iconadd2.gif" width="30" height="26" alt="ADD" border="0"></a></td>
		<%End if%>
                <td width="9%" bgcolor="#FFD553" align="center" height="8"><img src="../images/iconhelp2.gif" width="30" height="26" alt="HELP"></td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="15%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendor 
                  Code</font></b></font></td>
                <td width="25%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Vendor 
                  Name</font></b></font></td>
                <td width="35%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Address</font></b></font></td>
                <td width="15%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Tel No.</font></b></font></td>
                <td width="5%" bgcolor="#8B2718">&nbsp;</td>
                <td width="5%" bgcolor="#8B2718" align="right" valign="top"><img src="images/y05.jpg" width="12" height="12"></td>
              </tr>
              <%If  RS.EOF then 
       		Response.Write  " <tr bgcolor='#F1B3A7'> " & vbCrLf
		       Response.Write  " <td colspan='3' valign='top'>ไม่มีรายการการที่ค้นหา"
	       	Response.Write  "</td></tr> "& vbCrLf
		else
		RS.Absolutepage=PageNo
				RowCount=1
				Do While Not RS.EOF and RowCount<=PageSize 
				Vendcode=RS("svendcode")
				if keywords="" then 
					sSQL2="Select * from pvender where svendcode='" & Vendcode & "'" &_
					"Order By svendcode"
				else
					sSQL2 = "SELECT  *   FROM pvender Where svendcode='" & Vendcode & "'" &_
					" AND (sTname LIKE '%" &  keywords & "%' " &_
					" OR sEname Like '%" & keywords & "%' " &_
					" OR sVendcode Like '%" & keywords & "%' " &_
					" OR sSearchkey Like '%" & keywords & "%' )" &_
			 		" ORDER BY svendcode"
					End if

				Set Rs2 = Server.CreateObject("ADODB.RecordSet")
				Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
				If Not RS2.EOF then
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="2"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="15%" bgcolor="<%=Color%>" height="2"><%=RS2("svendcode")%>&nbsp;</td>
                <td width="25%" bgcolor="<%=Color%>" height="2"><a href="pvenderdesc.asp?svendcode=<%= Rs2("svendcode")%>"><%=RS2("sename")%></a></td>
                <td width="35%" bgcolor="<%=Color%>" height="2"><%=RS2("sEaddress")%>&nbsp;</td>
                <td width="15%" bgcolor="<%=Color%>" height="2"><%=RS2("sTelno")%>&nbsp;</td>
                <td width="5%" bgcolor="<%=Color%>" align="center" height="2">
		<%if Mid(Session("sAuthor"),7,3)="AED" then %>
				<a href='pvenderedt.asp?sprodcode=<%=Rs2("svendcode")%>&sexptype=<%=RS2("svendcode")%>&svendcode=<%=RS2("svendcode")%>'><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a></td>
		<%End if%>
                <td width="5%" bgcolor="<%=Color%>" align="center" height="2">
		<%if Mid(Session("sAuthor"),7,3)="AED" then %>
				<a href='pvenderdel.asp?sprodcode=<%=Rs2("svendcode")%>&sexptype=<%=RS2("svendcode")%>&svendcode=<%=RS2("svendcode")%>'><img src="../images/icondelete2.gif" width="30" height="26" alt="DELETE" border="0"></a> 
		<%End if%>
					<%Session("SearchVender")=keywords%>
					<%Session("pageno")=pageno%>
					<%Session("Maxpage")=maxpage%>
                </td>
              </tr>
              <%
				End if
				if Color="#F8D7D1" then
							Color="#F1B3A7"
				else
							Color="#F8D7D1"
				End if
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
        <%
if TotalPage > 1 then
%> 
        <form action= "<%=sScriptPath%>" method="Post">
          <tr> 
            <td colspan="4" align="center" bgcolor="#FFD553"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%
			if PageNo=1 then
					response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListVender=<%=sExptype%>&SearchVender=<%=keywords%>&page=Prev">Previous 
                    Page</a> <%
			End if
%> </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <%	
		If PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
		Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&ListVender=<%=sExptype%>&SearchVender=<%=keywords%>&page=Next">Next 
                    Page</a> <%End If%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) then
%> <Font color="Red"><%=i%></Font> 
                    <%Else%> 
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&ListVender=<%=sExptype%>&SearchVender=<%=keywords%>&MaxPage=<%=TotalPage%>"><%=i%></a> 
                    <%
		End if
	Next 
%></td>
                </tr>
              </table>
              <%End if%> 
            </td>
          </tr>
        </form>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="../datamanager/datamanager.asp">Data Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
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
