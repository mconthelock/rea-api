<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../List.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
Dim  Rs,sSQL,sSQL2,RS2,RS3,RS4,RS5,color
Dim sSearchpr,status,searchdate_day,searchdate_month,searchdate_year
Dim searchtodate_day,searchtodate_month,searchtodate_year,searchdate,searchtodate,spono
Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage,sdate,stodate
Dim sScriptPath,chkPr,inv,tax,dRcv
Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	sSearchpr=Ucase(Request("ssearchpr"))
	status=Request("status")
	searchdate_day=Request("searchdate_day")
	searchdate_month=Request("searchdate_month")
	searchdate_year=Request("searchdate_year")
	searchtodate_day=Request("searchtodate_day")
	searchtodate_month=Request("searchtodate_month")
	searchtodate_year=Request("searchtodate_year")
''	searchdate=searchdate_month & "/" & searchdate_day & "/" & searchdate_year
''	searchtodate=searchtodate_month & "/" & searchtodate_day & "/" & searchtodate_year
	searchdate= searchdate_day & "/" & searchdate_month & "/" & searchdate_year
	searchtodate= searchtodate_day & "/" & searchtodate_month & "/" & searchtodate_year
	if searchdate_day ="" then
		searchdate=Request("sdate")
		searchtodate=Request("stodate")	
	End if
	if cdate(searchdate) < cdate(searchtodate) then
		sdate=searchdate
		stodate=searchtodate
	else
		sdate=searchtodate
		stodate=searchdate
	End if  
	Set Rs = Server.CreateObject("ADODB.RecordSet")
'if cDate(searchdate) > cDate(searchtodate) then
'	sSQL="Select * from ppo where srefno like '%" & sSearchpr & "%'" &_
'	" and dpodate <=" & To_Date(sdate) &_
'	" and dpodate >=" & To_Date(stodate) &_
'	" Order By sRefno"
'else
	sSQL="Select  spono,srefno,sprno,sExptype,cstatus from ppo,ppr where ppo.sRefno like '%" & sSearchpr & "%'" &_
	" and ppo.dpodate >=" & To_Date(sdate) &_
	" and ppo.dpodate <=" & To_Date(stodate) &_
	" and ppo.srefno=ppr.sprno"&_
	" Order By sRefno"
'End if
'	sSQL="Select * from ppo where sRefno like '%" & sSearchpr & "%'" 
'response.write sSQL
	'response.write sSQL
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	page=Request("page")
	gPage=Request("PageNo")
	MaxPage=Request("MaxPage")
	TotalRows=Rs.RecordCount	
	RS.PageSize = 10
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
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
                <td colspan="7" bgcolor="#DB4B37" valign="top" height="8"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%">
                    <tr>
                      <td width="90%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Search P/R For P/O : From <%=Show_date(sdate)%> To <%=show_date(stodate)%></font></b></font></td>
	                  <td width="9%" align="center" height="8"><img src="../images/iconhelp2.gif" width="26" height="26" alt="HELP"></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>

              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/R No.</font></b></font></td>
				<td width="5%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Exp </font></b></font></td>
                <td width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O No.</font></b></font></td>
                <td width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Receive No.</font></b></font></td>
                <td width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Invoice No.</font></b></font></td>
                <td width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Tax No.</font></b></font></td>
              </tr>
<%if Not Rs.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize %>
<%
		spono=Rs("spono")
%>
              <tr valign="top" bgcolor="<%=Color%>"> 
                <td width="5%" bgcolor="<%=Color%>" height="2"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="15%" bgcolor="<%=Color%>" height="2">
				<%if chkPr <> RS("sRefno") then%>
					<a href="../workflow/pprdesc.asp?sprno=<%=RS("sRefno")%>">
					<%=RS("sRefno")%></a>
				<%else%>
					<img src="../images/dot.gif">
				<%end if%>
				</td>
                <td width="5%" bgcolor="<%=Color%>" height="25" align="center"><a href="pexpenstypedesc.asp?sExptype=<%=RS("sExptype")%>"><%=RS("sExpType")%></a></td>
                <td width="15%" bgcolor="<%=Color%>" height="2"><a href="../workflow/ppodesc.asp?spono=<%=RS("sPono")%>"><%if RS("cstatus")="9" then response.write "<Font color=Black><B><u>"%><%=RS("spono")%><%if RS("cstatus")="9" then response.write "</Font></u></b>"%></a>&nbsp;</td>
                <td width="15%" bgcolor="<%=Color%>"  height="2" valign="top"><%show_receive()%></td>
                <td width="15%" bgcolor="<%=Color%>"  height="2" valign="top"><%CinvTax()%><%IF (Trim(inv) <> Trim(tax))  and (Trim(inv)<>"" and  Trim(tax)<>"") then response.write "<font color='red'>*</font>"%><%show_Invoice()%></td>
                <td width="15%" bgcolor="<%=Color%>"  height="2" valign="top"><%IF (Trim(inv) <> Trim(tax)) and (Trim(inv)<>"" and  Trim(tax)<>"")  then response.write "<font color='red'>*</font>"%><%show_Tax()%></td>
              </tr>
<%
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
		RowCount=RowCount+1
		chkPr=RS("sRefno")
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
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&sSearchpr=<%=sSearchpr%>&sdate=<%=searchdate%>&stodate=<%=searchtodate%>">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&sSearchpr=<%=sSearchpr%>&sdate=<%=searchdate%>&stodate=<%=searchtodate%>">Next 
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
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&sSearchpr=<%=sSearchpr%>&sdate=<%=searchdate%>&stodate=<%=searchtodate%>"><%=i%></a> 
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
          <td width="2%" align="left" valign="bottom" height="17"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="middle" align="center" height="17"><br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="../workflow/workflow.asp">Workflow Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
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
</body>
</html>
<%
Sub show_Receive()
	Dim RS2,sSQL2
	sSQL2="Select sRecby,sRecno from preceive where spono='" & spono & "'" &_
	"order by sRecNo"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not Rs2.EOF then
		Do While Not RS2.EOF
			response.write "<a href='prectaxdesc.asp?sRecno="&RS2("sRecno")&"&EVENT=Show'>" &RS2("sRecby") & "</a>" & "<br>"
		RS2.MoveNext
		Loop
	End if
	RS2.Close
	Set Rs2=Nothing
End Sub

Sub get_Receive()
	Dim RS2,sSQL2
	sSQL2="Select sRecby,sRecno from preceive where spono='" & spono & "'" &_
	"order by sRecNo"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not Rs2.EOF then
		Do While Not RS2.EOF
			dRcv = RS2("sRecby")
		'	response.write "<a href='prectaxdesc.asp?sRecno="&RS2("sRecno")&"&EVENT=Show'>" &RS2("sRecby") & "</a>" & "<br>"
		RS2.MoveNext
		Loop
	End if
	RS2.Close
	Set Rs2=Nothing
End Sub

Sub CInvTax()
	Dim RS2,sSQL2
	sSQL2="Select sVendInv,sInvno from pInvoice where spono='" & spono & "'" &_
	"order by sRecNO"
'	response.write "RCV["
'	show_receive()
'	response.write "]"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not Rs2.EOF then
		inv=RS2("sVendInv")
	Else
		inv=""
	End if
	'response.write "INV=[" & inv & "]"
	RS2.Close
	'sSQL2="Select sVendTax from pTax where spono='" & spono & "'"
	sSQL2="Select sVendTax from pTax where spono='" & spono & "' order by sRecNo"
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not Rs2.EOF then
		tax=RS2("sVendtax")
	Else
		tax=""
	End if
	'response.write "TAX=[" & tax & "]"
	RS2.Close
	Set Rs2=Nothing
End Sub

Sub show_Invoice()
	Dim RS2,sSQL2
	sSQL2="Select sVendInv,sInvno from pInvoice where spono='" & spono & "'" &_
	"order by sRecNo"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not Rs2.EOF then
		Do While Not Rs2.Eof 
			response.write "<a href ='pInvoicedesc.asp?sInvno="&RS2("sInvno")&"&EVENT=Show'>"& RS2("sVendInv")&"</a>" &"<br>"
		Rs2.MoveNext
		Loop
	End if
	RS2.Close
	Set Rs2=Nothing
End Sub

Sub show_Tax()
	Dim RS2,sSQL2
	sSQL2="Select sVendTax,sTaxno from pTax where spono='" & spono & "'" &_
	"order by sRecNo"
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	if Not Rs2.EOF then
		Do While Not Rs2.Eof 
			response.write "<a href ='pTaxdesc.asp?sTaxno="&RS2("sTaxno")&"&EVENT=Show'>"&RS2("sVendTax") &"</a>" &"<br>"
		Rs2.MoveNext
		Loop
	End if
	RS2.Close
	Set Rs2=Nothing
End Sub
%>