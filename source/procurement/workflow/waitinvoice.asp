<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage, sSearch
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")  

	IF Request("Search")<>"" then
		if request("sSearch") <> "" then
			sSearch=" and sRecBy like '%"& Request("sSearch") & "%'"  
		end if
		
		if request("sPOSearch") <> "" then
			sSearch= sSearch & " and preceive.sPoNo like '%"& Request("sPOSearch") & "%'"  
		end if
	Else
		sSearch = ""
	End IF
		     	
if Mid(Session("sAuthor"),31,3)="AED" then
	sSQL = "SELECT  *  FROM preceive  where  cstatus <>'A' and cstatus <> 'I'" &_
	" ORDER BY spono,sRecby "
	'######################			add column tax no (Purchase key for Fin.	)		####################
	'######################			modify by Kallayanee		####################
	'######################			date 26/08/2008		####################
	sSQL = "SELECT preceive.*, ptax.svendtax  FROM preceive, ptax  where  cstatus <> 'A' and cstatus <> 'I' and preceive.spono=ptax.spono(+) "  _
	& " and preceive.srecno=ptax.srecno(+)" 
	'######################			date 18/08/2010		####################
	'######################			search invoice no or po no		####################		
	if sSearch <> "" then
		sSQL = sSQL & sSearch
	end if
	sSQL = sSQL & " ORDER BY preceive.spono, preceive.sRecby "

end if
'response.write sSQL
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
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
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Waiting For 
                        Invoice </font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="10%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Rec   No. </font></b></font></td>
                <td width="10%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Tax   No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O  No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/R  No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Date</font></b></font><font color="#FFFFFF" size="2"> 
                  </font></b></font></td>
                <td width="15%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"></font></b></font><font color="#FFFFFF" size="2"> 
                  </font></b></font></td>
              </tr>
              <%
if Not Rs.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="25"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="10%" bgcolor="<%=Color%>" height="25"><a href ='../workflow/prectaxdesc.asp?srecno=<%=RS("sRecno")%>&EVENT=Show'><%=RS("sRecBy")%></a>&nbsp;</td>
                <td width="10%" bgcolor="<%=Color%>" height="25"><%=RS("svendtax")%></td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href="ppodesc.asp?spono=<%=RS("spono")%>"><%=RS("sPono")%></a></td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><%=RS("sPrno")%></td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><%=RS("dRecDate")%>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25">
					<a href="../workflow/pInvoiceedt.asp?spono=<%=Rs("sPono")%>&sRecno=<%=RS("sRecno")%>">Add  Invoice</a>&nbsp;</td>
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

        <form action= "<%=sScriptPath%>" method="Post">
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
		  		  <tr bgcolor="#FFD553"><td colspan="3" align="center">
		  <%	if Mid(Session("sAuthor"),31,3) = "AED" then%>
              <br>INV NO: <input type="text" name="sSearch">
			  <Br>P/O NO: <input type="text" name="sPOSearch"><br>
              <input type="submit" name="search" value="Search">
			  <%End IF %>
		  </td></tr>
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
