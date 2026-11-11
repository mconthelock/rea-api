<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo

   Dim Rs,sSQL
	Dim sEmpno, sName, sSurname
	Dim color,sScriptPath
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Color="#F8D7D1"
	'suserid = Request("uid")
	'swelmsg  =Request.Cookies("welcome")
	'scorpcode = session("gbcorpcode")
	'dDate = Date()
	' Employee List 
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	sSQL = 	"SELECT * FROM vaEmployee order by sEmpno"
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
'	Case " <<  "
'			PageNo=1
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
'		Case " >>  "
'				PageNo=TotalPage
'		Case "Goto Page"
'				If Cint(gPage)>= Cint(TotalPage) then
'						PageNo=TotalPage
'				elseif gPage<=1 then
'						PageNo=1
'				else
'						PageNo=gPage+1
'				End if
'		Case " "
'				PageNo=1
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
          <td width="96%" align="left" valign="top" bgcolor="#FFD553" height="2"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td colspan="7" bgcolor="#DB4B37" valign="top" height="9"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">User 
                        Manager </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
                <td width="12%" bgcolor="#FFD553" align="center" height="8"><a href="pEmployeeEdit.asp"><img src="../images/iconadd2.gif" width="26" height="26" alt="add file" border="0"></a></td>
                <td width="7%" bgcolor="#FFD553" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="help"></td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="3%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="9%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Emp 
                  No. </font></b></font></td>
                <td width="39%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Name 
                  Surname</font></b></font></td>
                <td width="8%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Div</font></b></font></td>
                <td width="6%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Dep</font></b></font></td>
                <td width="8%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Sec</font></b></font></td>
                <td width="8%" bgcolor="#8B2718">&nbsp;</td>
                <td width="12%" bgcolor="#8B2718" align="right" valign="top"><img src="images/y05.jpg" width="12" height="12"></td>
              </tr>
              <%If  Rs.EOF Then %> 
              <tr bgcolor="<%=Color%>"> 
                <td width="3%" bgcolor="<%=Color%>" height="2">&nbsp;</td>
                <td width="9%" bgcolor="<%=Color%>" height="2"><%Response.write  "ยังไม่มีรายการของบุคคล"%></td>
                <td width="39%" bgcolor="<%=Color%>" align="center" height="2" valign="top">&nbsp;</td>
                <td bgcolor="<%=Color%>" align="center" height="2" ColSpan="6"><a href='pEmployeeDel.asp?sPoscode=<%=RS("sEmpno")%>'> 
                  </a> </td>
              </tr>
              <% else
RS.Absolutepage=PageNo
RowCount=1
Do While Not RS.EOF and RowCount<=PageSize

  %> 
              <tr bgcolor="<%=Color%>"> 
                <td width="3%" bgcolor="<%=Color%>" height="2"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="9%" bgcolor="<%=Color%>" height="3"><b><%=RS("sEmpno")%>&nbsp;</b></td>
                <td width="39%" bgcolor="<%=Color%>" height="3"><%=RS("sName") & "    "%> 
                  <%=RS("sSurname")%></td>
                <td width="8%" bgcolor="<%=Color%>" height="3"><%=RS("sDivcode")%>&nbsp;</td>
                <td width="6%" bgcolor="<%=Color%>" height="3"><%=RS("sDepcode")%>&nbsp;</td>
                <td width="8%" bgcolor="<%=Color%>" height="3"><%=RS("sSeccode")%>&nbsp;</td>
                <td width="8%" bgcolor="<%=Color%>" align="center" height="3" valign="top">aaa</td>
                <td width="12%" bgcolor="<%=Color%>" align="center" height="3"><A HREF='pEmployeeEdit.asp?sEmpno=<%=RS("sEmpno")%>'><IMG SRC="../images/iconedit2.gif" WIDTH="26" HEIGHT="26" BORDER="0" ALT="EDIT"></A></td>
				<td WIDTH="7%" ALIGN="CENTER"><a href='pEmployeeDel.asp?sEmpno=<%=RS("sEmpno")%>'><img src="../images/icondelete2.gif" width="26" height="26" alt="DELETE" border="0"> 
                  </a> </td>
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
End IF
%> 
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
<%if TotalPage > 1 then%>

			<tr><td colspan="3" align="center" bgcolor="#FFD553">
			      
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td align="center" bgcolor="#FFD553" width="2%">&nbsp;</td>
                <td align="center" bgcolor="#FFD553" width="5%">&nbsp;</td>
                <td align="center" bgcolor="#FFD553">&nbsp;
                <td align="center" bgcolor="#FFD553" width="5%">&nbsp;</td>
                <td align="center" bgcolor="#FFD553" width="2%">&nbsp;</td>
              </tr>
              <tr> 
                <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif" width="8" height="8"></td>
                <td align="center" bgcolor="#FFD553" width="5%"> <% if PageNo=1 then 
response.write "Previous"
else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev">Previous</a> 
                  <%end if%> 
                </td>
                <td align="center" bgcolor="#FFD553"> <%	
	For i=1 to TotalPage
		if Cint(PageNo)=Cint(i) then
%> 
                  <font color="Red"><%=i%></font> 
                  <%else %> 
                  <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxpage=<%=TotalPage%>"><%=i%></a> 
                  <%
		End if
	Next 
%> 
                <td align="center" bgcolor="#FFD553" width="5%"> <%
if PageNo=TotalPage Or PageNo=MaxPage then
	Response.write "Next"
else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next">Next</a> 
                  <% End if%> 
                </td>
                <td align="center" bgcolor="#FFD553" width="2%"><img src="images/dot.gif"></td>
              </tr>
            </table>
<%End if%>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../Pdivision/pdivisionall.asp"> 
            </a> <br>
            </font></td>
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
