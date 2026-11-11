<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  *  FROM pInvoice  " &_
	" ORDER BY ddate "
if Mid(Session("sAuthor"),31,3) = "AED" or Mid(Session("sAuthor"),31,3) = "VIW" then
	sSQL = "SELECT  * FROM pInvoice "  &_
	" ORDER BY ddate DESC"
'	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
else
	sSQL = "SELECT  distinct(pinvoice.sinvno),pinvoice.sVendInv,pinvoice.sPono,pinvoice.sRecno,pinvoice.ddate,preceive.recby FROM pInvoice,preceive,pprapp  where  pinvoice.srecno=preceive.srecno" &_
	" and preceive.sprno=pprapp.sprno" &_
	" and pprapp.sEmpno='" & Session("sEmpno") & "'" &_
	" ORDER BY ddate DESC"
'	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
End if
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
            WORKFLOW MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td colspan="6" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Finished 
                        Invoice </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="6%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Invoice 
                  No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Rec 
                  No.</font></b></font><font color="#FFFFFF" size="2"> </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O 
                  No.</font></b></font></b></font><font color="#FFFFFF" size="2"> 
                  </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Date</font></b></font></b></font><font color="#FFFFFF" size="2"> 
                  </font></b></font></td>
                <td width="14%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><img src="../images/dot.gif"></font></b></font></b></font><font color="#FFFFFF" size="2"> 
              </tr>
              <%
if Not Rs.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="6%" bgcolor="<%=Color%>" height="25"><a href ='../workflow/pInvoicedesc.asp?sInvno=<%=RS("sInvno")%>&EVENT=Show'><%=((PageNo-1)*RS.PageSize)+RowCount%></a>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href ='../workflow/pInvoicedesc.asp?sInvno=<%=RS("sInvno")%>&EVENT=Show'><%=RS("sVendInv")%></a>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href="prectaxdesc.asp?sRecno=<%=RS("sRecno")%>"><%Show_Rec(RS("sRecno"))%></a></td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href="ppodesc.asp?spono=<%=RS("spono")%>"><%=RS("sPono")%></a>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><%=RS("ddate")%>&nbsp;</td>
                <td width="14%" bgcolor="<%=Color%>" height="25" align="Center"><img src="../images/dot.gif">
<%if Mid(Session("sAuthor"),31,3) = "AED" then%>
				<a href="../workflow/pInvoiceedt.asp?sInvno=<%=RS("sInvno")%>&status=pinvedt">
				<img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a>
				<a href="reinvoice.asp?sinvno=<%=RS("sInvno")%>">
				<img src="../images/icondelete2.gif" width="26" height="26" alt="DELETE" border="0"></a>&nbsp;
<%End if%>
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
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../datamanager/images/dot.gif"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%
			If PageNo=1 then
				response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next">Next 
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
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>"><%=i%></a> 
                    <%
	End if
	Next %></td>
                </tr>
              </table>
              <%End if%> 
            </td>
          </tr>
        </form>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="Workflow.asp">Workflow Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
            </font></td>
          <td width="2%" align="right" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
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
<%
Sub Show_Rec(sRecno)
	Dim Rs2,sSQL
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL="Select sRecby from preceive where srecno='" & sRecno & "'"
 	Rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF Not Rs2.Eof then
		Response.write RS2("sRecby")
	End if
	RS2.Close
	Set RS2=Nothing
End Sub
%>
