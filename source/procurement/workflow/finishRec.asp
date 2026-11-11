<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,chkInv,sRecno,spono
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
 if Mid(Session("sAuthor"),37,3) = "AED" then 
	
if Request("PONO") = "" then
sSQL = "SELECT distinct(preceive.srecno),preceive.sRecby,preceive.drecdate,preceive.sprno,preceive.spono,ppr.sExptype  FROM pReceive,ppr,pprapp  where ppr.cpo <>'4'" &_
" and preceive.sprno=ppr.sprno" &_
" and ppr.sprno=pprapp.sprno" &_
" ORDER BY pReceive.spono,pReceive.sRecby"
else
sSQL = "SELECT distinct(preceive.srecno),preceive.sRecby,preceive.drecdate,preceive.sprno,preceive.spono,ppr.sExptype  FROM pReceive,ppr,pprapp  where ppr.cpo <>'4'" &_
" and preceive.sprno=ppr.sprno" &_
" and ppr.sprno=pprapp.sprno" &_
" and pReceive.spono like '%"&trim(Request("PONO"))&"%'" &_
" ORDER BY pReceive.spono,pReceive.sRecby"
end if
else 
	if Request("PONO") = "" then
		sSQL = "SELECT distinct(preceive.srecno),preceive.sRecby,preceive.drecdate,preceive.sprno,preceive.spono,ppr.sExptype  FROM pReceive,ppr,pprapp  where ppr.cpo <>'4'" &_
		" and preceive.sprno=ppr.sprno" &_
		" and ppr.sprno=pprapp.sprno" &_
		" and (pprapp.sEmpno='" & Session("sEmpno") & "'" &_
		" Or ppr.sKey='" & Session("sEmpno") & "')"  &_
		" ORDER BY pReceive.spono desc,pReceive.sRecby"
	else
			sSQL = "SELECT distinct(preceive.srecno),preceive.sRecby,preceive.drecdate,preceive.sprno,preceive.spono,ppr.sExptype  FROM pReceive,ppr,pprapp  where ppr.cpo <>'4'" &_
		" and preceive.sprno=ppr.sprno" &_
		" and ppr.sprno=pprapp.sprno" &_
		" and (pprapp.sEmpno='" & Session("sEmpno") & "'" &_
		" Or ppr.sKey='" & Session("sEmpno") & "') and pReceive.sPoNo like '%"& Request("sPOSearch") & "%'"   &_
		" ORDER BY pReceive.spono desc,pReceive.sRecby"
	end if 

end if
	'response.write sSQl
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
                <td colspan="7" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr> 
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Finish 
                        Receiver </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="6%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Rec 
                  No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O 
                  No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/R 
                  No. </font></b></font></td>
				  <td width="5%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Exp </font></b></font></td>
                <td width="15%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Date</font></b></font><font color="#FFFFFF" size="2"> 
                  </font></b></font></td>
                <td width="14%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> <img src="../images/dot.gif">
                  </font></b></font></td>
              </tr>
              <%
if Not Rs.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="6%" bgcolor="<%=Color%>" height="25"><a href ='../workflow/prectaxdesc.asp?srecno=<%=RS("sRecno")%>&EVENT=Show'><%=((PageNo-1)*RS.PageSize)+RowCount%></a>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href ='../workflow/prectaxdesc.asp?srecno=<%=RS("sRecno")%>&EVENT=Show'><%=RS("sRecBy")%></a>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href="ppodesc.asp?spono=<%=RS("spono")%>"><%=RS("sPono")%></a></td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href ='showprdesc.asp?sprno=<%=RS("sPrno")%>'><%=RS("sPrno")%></a></td>
                <td width="5%" bgcolor="<%=Color%>" height="25" align="center"><%=RS("sExpType")%></td>
				<td width="15%" bgcolor="<%=Color%>" height="25"><%=RS("dRecDate")%>&nbsp;</td>
                <td width="14%" bgcolor="<%=Color%>" height="25" align="center"><img src="../images/dot.gif">
<%	sRecno=RS("sRecno")
	chkInvoice()%>
<% if chkInv=0 and Mid(Session("sAuthor"),37,3) = "AED"  then%>
	<a href="preceiveedt2.asp?srecno=<%=RS("sRecno")%>&spono=<%=RS("spono")%>"><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a>&nbsp;
	<a href="preceivedel.asp?srecno=<%=RS("sRecno")%>&sprno=<%=RS("sprno")%>&spono=<%=RS("spono")%>"><img src="../images/icondelete2.gif" width="26" height="26" alt="DELETE" border="0"></a>&nbsp;
<% End if%>
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
		<tr><td colspan="4" align="center" bgcolor="#FFD553"><br/><br/><br/></td></tr>
		<tr><td colspan="4" align="center" bgcolor="#FFD553">   <form action= "<%=sScriptPath%>" method="Post">P/O NO: <input type="text" name="PONO" id="PONO" value=""/>&nbsp;&nbsp;<input type="submit" name="search" id="search" value="Search"/></form></td></tr>

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
            | <a href="workflow.asp">Workflow Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
            </font></td>
          <td width="2%" align="right" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
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
</body>
</html>
<%
Sub chkinvoice()
    Dim  Rs2,sSQL2,amount
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select * from pInvoice where sRecno='" & sRecno & "'" 
	Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if Not Rs2.EOF then
		chkInv=1
	else
		chkInv=0
	End if
	Rs2.Close
	Set RS2=nothing
 End Sub
%>