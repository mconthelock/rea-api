<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../List.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
		
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       		
if Request("PONO") = "" then
	sSQL="Select distinct(ppo.spono),ppr.sprno,ppr.skey,ppr.sRequester,ppr.dprdate,sExptype  from ppr,pprapp,ppo  where  ppr.cpo='2'" &_
	" and ppr.sprno=pprapp.sprno" &_
	" and ppr.sprno=ppo.srefno" &_
	" and ppo.cStatus<>'9'" &_
	" and ppo.cstatus='2'" &_
	" Order By ppo.spono"
else
	sSQL="Select distinct(ppo.spono),ppr.sprno,ppr.skey,ppr.sRequester,ppr.dprdate,sExptype  from ppr,pprapp,ppo  where  ppr.cpo='2'" &_
	" and ppr.sprno=pprapp.sprno" &_
	" and ppr.sprno=ppo.srefno" &_
	" and ppo.cStatus<>'9'" &_
	" and ppo.cstatus='2'" &_
	" and ppo.spono like '%"&trim(Request("PONO"))&"%'" &_
	" Order By ppo.spono"
end if


'response.write sSQL
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
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Waiting  
                        For Receive</font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37">  
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/R No. </font></b></font></td>
                <td width="5%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Exp </font></b></font></td>
				<td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Key</font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Requester</font></b></font></td>
                <td width="15%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Date</font></b></font></td>
              </tr>
              <%
if Not Rs.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize
%>  
              <tr bgcolor="<%=Color%>">  
                <td width="5%" bgcolor="<%=Color%>" height="25"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href ='../workflow/pReceiveedt.asp?sprno=<%=RS("sPrno")%>&spono=<%=RS("sPono")%>'><%=RS("sPono")%></a>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href ='showprdesc.asp?sprno=<%=RS("sPrno")%>'><%=RS("sPrno")%></a></td>
				<td width="5%" bgcolor="<%=Color%>" height="25" align="center"><%=RS("sExpType")%></td>
                <td width="20%" bgcolor="<%=Color%>" height="25">
				<font size="1">		
				<%Show_Empno(RS("sKey"))%></td>
                <td width="20%" bgcolor="<%=Color%>" height="25">
				<font size="1">		
				<%Show_Empno(RS("sRequester"))%>&nbsp;</td>
                <td width="15%" bgcolor="<%=Color%>" height="25"><%=RS("dPrdate")%>&nbsp;</td>
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
            | <a href="Workflow.asp">Workflow Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
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
