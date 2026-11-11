<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../list.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
Dim  Rs,sSQL,sSQL2,RS2,RS3,RS4,RS5,color
Dim sSearchpr,status,searchdate_day,searchdate_month,searchdate_year
Dim searchtodate_day,searchtodate_month,searchtodate_year,searchdate,searchtodate
Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,sdate,stodate
Dim sScriptPath
Color="#F8D7D1"
	If Request("Search")<>"" then
		sSearchpr=Ucase(Request("ssearch"))
	Else
		sSearchPr="AMEC"
	End IF

	Set Rs = Server.CreateObject("ADODB.RecordSet")

		sSQL="Select * from ppr where  sPrno like '%" & sSearchPr & "%' and sExptype='" & Request("sExptype") & "'" &_
	" and (sKey  =  '" & Session("sEmpNo") & "'" &_
	" or     sRequester= '" & Session("sEmpNo") & "')" &_
	" and cApprove<>'9'" &_	
	" Order By sprno"
'response.write sSQl

 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
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
<title>PPR</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="POST" action="searchprcopy.asp">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
    <td width="63%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="1" height="1"></td>
    <td width="36%" height="2" align="center"><img src="../images/spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td width="1%" height="2"><img src="../images/spacer.gif" width="1" height="1"></td>
    <td colspan="2" height="2" align="center" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">...</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td colspan="5" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
 
				  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr>
                      <td width="90%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Search P/R : <input type="text" name="ssearch">
					  <input type="hidden" name="sExptype" value="<%=Request("sExptype")%>">
					  <input type="hidden" name="sNewPrno" value="<%=Request("sNewPrno")%>">
					  <input type="submit" name="search" value="Search"></font></b></font></td>
	                  <td width="9%" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="HELP"></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>

                </td>

              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="3%" bgcolor="#8B2718" HEIGHT="5"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="12%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/R No.</font></b></font></td>
                <td width="35%" bgcolor="#8B2718"  align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Key/Requester</font></b></font></td>
                <td width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Date</font></b></font></td>
                <td width="3%" bgcolor="#8B2718" align="right" valign="top"><img src="../images/dot.gif" width="1" height="1"></td>
              </tr>
<% if Not Rs.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize  %>
              <tr bgcolor="<%=Color%>"> 
                <td width="3%" bgcolor="<%=Color%>" height="10"> 
                  <%=((PageNo-1)*RS.PageSize)+RowCount%>.&nbsp;
                </td>
                <td width="12%" bgcolor="<%=Color%>" height="2"><a href="../workflow/showprdesc.asp?sprno=<%=RS("sPrno")%>"><%=RS("sprno")%></a>&nbsp;</td>
                <td width="35%" bgcolor="<%=Color%>" height="2"><font size="1">
				<%Show_Empno(Rs("sKey"))%>
				<font color="teal">
				/
				</font>
		<%if RS("cPrtype")="1" then response.write "<font color='red'><b>"%>
				<%Show_Empno(RS("srequester"))%></td>
                <td width="10%" bgcolor="<%=Color%>"  height="2" valign="top"><%=Rs("dprdate")%></td>
		<%if RS("cPrtype")="1" then response.write "</b></font>"%>
                  <td width="3%" bgcolor="<%=Color%>" align="center" height="2"><a href="addprcopy.asp?sprno=<%=RS("sprno")%>&sNewPrno=<%=Request("sNewPrno")%>">Add</a><img src="../images/dot.gif" width="8" height="8"></td>
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
%>
            </table>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <%if TotalPage > 1 then%> 

          <tr> 
            <td colspan="4" align="center" bgcolor="#FFD553"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../datamanager/images/dot.gif"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%
			If PageNo=1 then
				response.write " "
			Else
%> <A Href="searchprcopy.asp?PageNo=<%=PageNo%>&page=Prev&status=<%=status%>&sSearchpr=<%=sSearchpr%>&sdate=<%=searchdate%>&stodate=<%=searchtodate%>&sExptype=<%=Request("sExptype")%>&sNewPrno=<%=Request("sNewPrno")%>&Search=s">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="searchprcopy.asp?PageNo=<%=PageNo%>&page=Next&status=<%=status%>&sSearchpr=<%=sSearchpr%>&sdate=<%=searchdate%>&stodate=<%=searchtodate%>&sExptype=<%=Request("sExptype")%>&sNewPrno=<%=Request("sNewPrno")%>&Search=s">Next 
                    Page</a> <%End if%> 
                  </td>
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../datamanager/images/dot.gif"></td>
                </tr>
                <tr> 
                  <td align="center" bgcolor="#FFD553" colspan="4"> <%
	For i=1 to TotalPage
		If Cint(PageNo)=Cint(i) Then
%> <Font color="Red"><%=i%></font> 
                    <%else%> 
                    <A Href="searchprcopy.asp?PageNo=<%=i%>&status=<%=status%>&maxPage=<%=TotalPage%>&sSearchpr=<%=sSearchpr%>&sdate=<%=searchdate%>&stodate=<%=searchtodate%>&sExptype=<%=Request("sExptype")%>&sNewPrno=<%=Request("sNewPrno")%>&Search=s"><%=i%></a> 
                    <%
	End if
	Next %></td>
                </tr>
              </table>
              <%End if%> 
            </td>
          </tr>

        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom" height="17"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center" height="17"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="../datamanager/datamanager.asp">Data Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
            </font></td>
          <td width="2%" align="right" valign="bottom" height="17"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
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
</form>
</body>
</html>
