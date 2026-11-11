<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo

   Dim Rs,sSQL
	Dim sEmpno, sName, sSurname
	Dim unEmpCount, PageRowCount
	Dim color,sScriptPath,SQL
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,keywords,sposcode,sDivcode,sDepcode,sSeccode,Orderby,ob
	Color="#F8D7D1"
if Request("status")="del" then
'	page=Session("page")
	gPage=Session("PageNo")
	MaxPage=Session("MaxPage")
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	keywords = UCase(TRIM(Session( "SearchUser" ) ))
	sDivcode=Session("ListDivision")
	sDepcode=Session("ListDepartment")
	sSeccode=Session("ListSection")
	sposcode=Session("ListPosition")
	ob=Session("ListOrderby")
else
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	keywords = UCase(TRIM( Request( "SearchUser" ) ))
	sDivcode=Request("ListDivision")
	sDepcode=Request("ListDepartment")
	sSeccode=Request("ListSection")
	sposcode=Request("ListPosition")
	ob=Request("ListOrderby")
End if
if ob="1" then 
		orderby="sTname"
elseif ob="2" then 
		orderby="sName"
elseif ob="3" then 
		orderby="sDivcode"
elseif ob="4" then 
		orderby="sDepcode"
elseif ob="5" then 
		orderby="sSeccode"
elseif ob="6" then 
		orderby="sPoscode"
else 
		orderby="sEmpno"
end if
if keywords="" then
	if  (sposcode="0" Or sposcode="00") and (sDivcode="0" Or sDivcode="00") and (sDepcode="0" Or sDepcode="00") and (sSeccode="0" Or sSeccode="00") then
			sSQL = 	"SELECT * FROM aEmployee order by " & orderby 
	elseif  (sDepcode ="0" Or sDepcode ="00")  and (sSeccode ="0" Or sSeccode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sSeccode ="0" Or sSeccode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sDepcode ="0" Or sDepcode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sSeccode='" & sSeccode & "'" &_
	 		" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sDepcode ="0" Or sDepcode ="00") and (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sPoscode='" & sPoscode & "'" &_
	 		" ORDER BY " & orderby
	elseif  (sSeccode ="0" Or sSeccode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
		"and sDepcode='" & sDepcode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
		"and sSeccode='" & sSeccode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sDepcode ="0" Or sDepcode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sSeccode='" & sSeccode & "'" &_
		"and sPoscode='" & sPoscode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDepcode ="0" Or sDepcode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
		"and sSeccode='" & sSeccode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDepcode ="0" Or sDepcode ="00") and (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
		"and sPoscode='" & sPoscode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
		"and sPoscode='" & sPoscode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
		"and sSeccode='" & sSeccode & "'" &_
		"and sPoscode='" & sPoscode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sDepcode ="0" Or sDepcode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
		"and sSeccode='" & sSeccode & "'" &_
		"and sPoscode='" & sPoscode & "'" &_
 			" ORDER BY " & orderby
	elseif  (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
		"and sDepcode='" & sDepcode & "'" &_
		"and sPoscode='" & sPoscode & "'" &_
 		" ORDER BY " & orderby
	elseif  (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
		"and sDepcode='" & sDepcode & "'" &_
		"and sSeccode='" & sSeccode & "'" &_
 			" ORDER BY " & orderby
	else 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sDepcode='" & sDepcode & "'" &_
			"and sSeccode='" & sSeccode & "'" &_
			" and sPoscode='" & sPoscode & "'" &_
 			" ORDER BY " & orderby
	end if
else 
	if  (sPoscode ="0" Or sPoscode ="00") and (sDivcode ="0" Or sDivcode ="00") and (sDepcode ="0" Or sDepcode ="00") and (sSeccode ="0" Or sSeccode ="00") then

	sSQL = "SELECT  * FROM aemployee Where sname LIKE '%" &  keywords & "%' " &_
		" OR stname Like '%" & keywords & "%' " &_
		" OR sempno Like '%" & keywords & "%' " &_
 		" ORDER BY " & orderby
	elseif  (sDepcode ="0" Or sDepcode ="00") and (sSeccode ="0" Or sSeccode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sSeccode ="0" Or sSeccode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sDepcode ="0" Or sDepcode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sSeccode='" & sSeccode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
	 		" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sDepcode ="0" Or sDepcode ="00") and (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
	 		" ORDER BY " & orderby
	elseif  (sSeccode ="0" Or sSeccode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sDepcode='" & sDepcode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode="0" Or sDivcode="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
			"and sSeccode='" & sSeccode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") and (sDepcode ="0" Or sDepcode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sSeccode='" & sSeccode & "'" &_
			"and sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sDepcode ="0" Or sDepcode ="00") and (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sSeccode='" & sSeccode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sDepcode ="0" Or sDepcode ="00") and (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00")  and (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
			"and sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sDivcode ="0" Or sDivcode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDepcode='" & sDepcode & "'" &_
			"and sSeccode='" & sSeccode & "'" &_
			"and sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif (sDepcode ="0" Or sDepcode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sSeccode='" & sSeccode & "'" &_
			"and sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_ 	
			" ORDER BY " & orderby
	elseif  (sSeccode ="0" Or sSeccode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sDepcode='" & sDepcode & "'" &_
			"and sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	elseif  (sPoscode ="0" Or sPoscode ="00") then 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sDepcode='" & sDepcode & "'" &_
			"and sSeccode='" & sSeccode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	else 
			sSQL = "SELECT  * FROM aEmployee Where sDivcode='" & sDivcode & "'" &_
			"and sDepcode='" & sDepcode & "'" &_
			"and sSeccode='" & sSeccode & "'" &_
			" and sPoscode='" & sPoscode & "'" &_
			" AND (sname LIKE '%" &  keywords & "%' " &_
			" OR sempno Like '%" & keywords & "%' " &_
			" OR stname Like '%" & keywords & "%' )" &_
 			" ORDER BY " & orderby
	end if
End if

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
<title>User Manager</title>
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
          <td align="right" valign="top" width="2%" height="2"><img src="images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td colspan="7" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">User 
                        Manager </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
                <td width="6%" bgcolor="#FFD553" align="center" height="8"><%
	if Mid(Session("sAuthor"),4,1) = "A" OR Mid(Session("sAuthor"),19,1)="A" Then
%> <A HREF="../datamanager/pEmployeeEdit.asp"><IMG SRC="../images/iconadd2.gif" WIDTH="26" HEIGHT="26" BORDER="0" ALT="Add"></A> 
                  <% End if%></td>
                <td width="8%" bgcolor="#FFD553" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="help" BORDER="0"></td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="3%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="14%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Emp 
                  No. </font></b></font></td>
                <td width="46%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">&nbsp;&nbsp;Name 
                  Surname</font></b></font></td>
                <td width="6%" bgcolor="#8B2718"> 
                  <DIV ALIGN="CENTER"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Div</font></b></font></DIV>
                </td>
                <td width="5%" bgcolor="#8B2718"> 
                  <DIV ALIGN="CENTER"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Dep</font></b></font></DIV>
                </td>
                <td width="6%" bgcolor="#8B2718"> 
                  <DIV ALIGN="CENTER"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Sec</font></b></font></DIV>
                </td>
                <td width="6%" bgcolor="#8B2718">&nbsp;<%=(PageNo * PageSize) - PageSize%></td>
                <td width="6%" bgcolor="#8B2718">&nbsp;</td>
                <td width="8%" bgcolor="#8B2718" align="right" valign="top"><img src="images/y05.jpg" width="12" height="12"></td>
              </tr>
              <%If  Rs.EOF Then %> 
              <tr bgcolor="<%=Color%>"> 
                <td width="3%" bgcolor="<%=Color%>" height="2">&nbsp;</td>
                <td width="14%" bgcolor="<%=Color%>" height="2"><%Response.write  "ไม่มีรายการ"%></td>
                <td width="46%" bgcolor="<%=Color%>" align="center" height="2" valign="top">&nbsp;</td>
                <td width="6%" bgcolor="<%=Color%>" align="center" height="2"> 
                </td>
                <td bgcolor="<%=Color%>" align="center" height="2" Colspan="5"> 
                </td>
              </tr>
              <% else
						RS.Absolutepage=1 'PageNo
						PageRowCount=1
						RowCount = 1
						
						Do While Not RS.EOF and RowCount <= (PageNo * PageSize) - PageSize
							if rs("cStatus") <> "0" then
								PageRowCount = PageRowCount + 1
							end if
							
							rowCount = rowCount + 1
							
							rs.moveNext
						loop

						Do While Not RS.EOF and RowCount <= PageNo * PageSize
							'########## ADD BY Dermpan, 23-DEC-02 ##########
										  if rs("cStatus") = "0" then
											color = "gray"
							elseif Color="#F8D7D1" then
									Color="#F1B3A7"
							else
									Color="#F8D7D1"
							end if
							'########## END ADD BY Dermpan, 23-DEC-02 ##########
	
					  %> 
              <tr bgcolor="<%=Color%>"> 
                <!--<td width="3%" bgcolor="<%=Color%>" height="20"><%if rs("cStatus") = "0" then response.write "#" else response.write ((PageNo-1)*RS.PageSize)+PageRowCount end if%>&nbsp;</td>-->
                <td width="3%" bgcolor="<%=Color%>" height="20"><%if rs("cStatus") = "0" then response.write "#" else response.write PageRowCount end if%>&nbsp;</td>
                <td width="14%" bgcolor="<%=Color%>" height="20"> <b> <%	if Mid(Session("sAuthor"),1,3) = "999" Then%> 
                <%=RS("sEmpno")%></b> <%else%> <B> <%=RS("sEmpno")%></B> <% End if%> 
                </td>
                <td width="46%" bgcolor="<%=Color%>" height="20"> <%	if Mid(Session("sAuthor"),1,3) <> "999" Then%> 
<a href ="../datamanager/pEmployeedesc.asp?sEmpno=<%=RS("sEmpno")%>">
                  <%=RS("sName")%></a><% else %>  
<a href ="../datamanager/pEmployeedesc.asp?sEmpno=<%=RS("sEmpno")%>">
                  <%=RS("sName")%></a><%end if%> </td>
                <td width="6%" bgcolor="<%=Color%>" height="20"><%=RS("sDivcode")%>&nbsp;</td>
                <td width="5%" bgcolor="<%=Color%>" height="20"><%=RS("sDepcode")%>&nbsp;</td>
                <td width="6%" bgcolor="<%=Color%>" height="20"><%=RS("sSeccode")%>&nbsp;</td>
                <td width="6%" bgcolor="<%=Color%>" align="center" height="20"><a href="../Security/aAuthorizeedt.asp?sEmpno=<%=RS("sEmpno")%>"><%
	if  Mid(Session("sAuthor"),1,3)="999" Then
%><IMG SRC="../images/security26.gif" WIDTH="26" HEIGHT="26" BORDER="0" ALT="Authroize Edit"><% End if%></a><img src="../images/dot.gif"></td>
                <td width="6%" bgcolor="<%=Color%>" align="center" height="20"><a href='pEmployeeEdit.asp?sEmpno=<%=RS("sEmpno")%>'><%
	if Mid(Session("sAuthor"),19,3)="AED" OR Mid(Session("sAuthor"),1,3)="999"  Then
%><img src="../images/iconedit2.gif" width="26" height="26" border="0" alt="EDIT"><% End if%></a></td>
                <td width="8%" bgcolor="<%=Color%>" align="center" height="20"><a href='pEmployeeDel.asp?sEmpno=<%=RS("sEmpno")%>'>
<%Session("PageNo")=PageNo%>
<%Session("searchUser")=keywords%>
<%Session("ListPosition")=sposcode%>
<%Session("ListDivision")=sDivcode%>
<%Session("ListDepartment")=sDepcode%>
<%Session("ListSection")=sSeccode%>
<%Session("ListOrderby")=ob%>
<%
	if Mid(Session("sAuthor"),19,3)="AED"  OR Mid(Session("sAuthor"),1,3)="999"  Then
%><img src="../images/icondelete2.gif" width="26" height="26" alt="DELETE" border="0"><% End if%> 
                  </a> </td>
              </tr>
              <%
'			  if Color="#F8D7D1" then
'		Color="#F1B3A7"
'else
'		Color="#F8D7D1"
'end if

'########## MODIFY BY Dermpan, 23-DEC-02 ##########
'RowCount=RowCount+1
'Modify by Dermpan, 23-DEC-02
RowCount = RowCount + 1

if rs("cStatus") <> "0" then
	PageRowCount = PageRowCount + 1
else
	unEmpCount = unEmpCount + 1
end if
'########## END MODIFY BY Dermpan, 23-DEC-02 ##########

 RS.MoveNext
Loop
End IF
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <%if TotalPage > 1 then%> 
        <tr> 
          <td colspan="3" align="center" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
                <td align="left" bgcolor="#FFD553" width="45%"> 
<%
		 if PageNo=1 then 
				response.write " "
		else
%>
 <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&searchUser=<%=keywords%>&ListPosition=<%=sposcode%>&ListDivision=<%=sDivcode%>&ListDepartment=<%=sDepcode%>&ListSection=<%=sSeccode%>&ListOrderby=<%=ob%>">Previous Page</a> 
  <%
		end if
%>
                </td>
                <td align="right" bgcolor="#FFD553" width="45%"> 
<%
if PageNo=TotalPage Or PageNo=MaxPage then
	Response.write " "
else
%>
 <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&searchUser=<%=keywords%>&ListPosition=<%=sposcode%>&ListDivision=<%=sDivcode%>&ListDepartment=<%=sDepcode%>&ListSection=<%=sSeccode%>&ListOrderby=<%=ob%>">Next Page</a> 
 <% 
End if
%> 
                </td>
                <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
              </tr>
              <tr>
                <td align="center" bgcolor="#FFD553" width="2%">&nbsp;</td>
                <td align="center" bgcolor="#FFD553" width="5%">&nbsp;</td>
                <td align="center" bgcolor="#FFD553">&nbsp;
                <td align="center" bgcolor="#FFD553" width="2%">&nbsp;</td>

              </tr>
<tr>
                <td align="center" bgcolor="#FFD553" colspan="6">
 <%	
	For i=1 to TotalPage
		if Cint(PageNo)=Cint(i) then
%>
 <font color="Red"><%=i%></font> 
                  <%else %> 
                  <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxpage=<%=TotalPage%>&searchUser=<%=keywords%>&ListPosition=<%=sposcode%>&ListDivision=<%=sDivcode%>&ListDepartment=<%=sDepcode%>&ListSection=<%=sSeccode%>&ListOrderby=<%=ob%>"><%=i%></a> 
                  <%
		End if
	Next 
%> 
</tr>
            </table>
            <%End if%> 
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
            </font><img src="../images/dot.gif"></td>
          <td width="2%" align="right" valign="bottom"><img src="images/y04.jpg" width="12" height="12"></td>
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
