<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	
    Dim  Rs,sSQL
	Dim sScriptPath,svendcode,Color,chkRec,spono,chkInv,chkRec2,sSearch
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
	Color="#F8D7D1"
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	IF Request("sSearch")<>"" then
		sSearch=" and ppo.spono like '%" & Request("sSearch") & "%'"
	Else
		sSearch=" and (ppo.cTax <>'2' or ppo.cInvoice<>'2' or ppo.cTax is NUll or ppo.cInvoice is Null) " 
	End IF
sSQL = "SELECT  distinct(ppo.spono),ppo.dpodate,ppr.sprno,ppr.sExptype,ppr.sCashReq FROM pPO,ppr,pprapp where (ppo.cStatus='2' or ppo.cStatus='3') " &_
	" and ppo.srefno=ppr.sprno" &_
	" and ppr.cpo <> '1'" & sSearch &_
	" and ppr.sprno=pprapp.sprno" &_
	" and (pprapp.sEmpno='" & Session("sEmpno") &_
	"' or pprapp.sEmpnoapp='" & Session("sEmpno") &_
	"' or ppr.sKey='" & Session("sEmpno") &_
	"') " &_
	" ORDER BY sPono DESC"
'	response.write sSQL
'		sSQL = "SELECT distinct (pprapp.sprno),ppo.srefno,ppo.spono,ppo.dpodate  FROM  ppo,pprapp WHERE  ppo.sRefno=pprapp.sPrno" &_
'		" and (pprapp.sEmpno='" & Session("sEmpno") &_
'		"' or pprapp.sEmpnoapp='" & Session("sEmpno") &_
'		"') and cStatus<>'2' and cStatus<>'3'"
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
                      <td width="99%" height="24"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O 
                        Printed</font></b></font></td>
                      <td width="1%" align="right" valign="top" height="24"><img src="../images/y06.jpg" width="12" height="12"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td width="6%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">No.</font></b></font></td>
                <td width="25%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/O 
                  No. </font></b></font></td>
                <td width="20%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">P/R 
                  No. </font></b></font></td>
				  <td width="5%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Exp </font></b></font></td>
                <td width="23%" bgcolor="#8B2718" align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Date</font></b></font></td>
                <td width="21%" bgcolor="#8B2718" align="center" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font><img src="../images/dot.gif"></td>
              </tr>
              <%
if Not Rs.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize
%> 
              <tr bgcolor="<%=Color%>"> 
                <td width="6%" bgcolor="<%=Color%>" height="25"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;</td>
                <td width="25%" bgcolor="<%=Color%>" height="25"><a href ='../workflow/ppodesc.asp?spono=<%=RS("sPono")%>'><%IF RS("sCashReq")<>"" then response.write "<Font color=#CC00CC>"%><%=RS("sPono")%></a>&nbsp;</td>
                <td width="20%" bgcolor="<%=Color%>" height="25"><a href ='showprdesc.asp?sprno=<%=RS("sPrno")%>'><%=RS("sPrno")%></a></td>
                 <td width="5%" bgcolor="<%=Color%>" height="25" align="center"><%=RS("sExpType")%></td>
				<td width="23%" bgcolor="<%=Color%>" height="25"><%=RS("dPodate")%>&nbsp;</td>
                <td width="5%" bgcolor="<%=Color%>" height="25" align="center">
<%
		if Mid(Session("sAuthor"),22,3)="AED" then %>
			<a href="ppoprintedt.asp?spono=<%=RS("sPono")%>"><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a>
		<%else %>
			<img src="../images/dot.gif">
		<%End if%>

</td>
<!--                <td width="16%" bgcolor="<%=Color%>" height="25"> 
<%
		spono=RS("spono")
'		chkReceive2()
'		chkInvoice()
		if Mid(Session("sAuthor"),22,3)="AED" then
			if chkInv="0" and chkrec2="0" then
%>
			<a href="cancelpo.asp?spono=<%=RS("spono")%>&sprno=<%=Rs("sprno")%>"></a>
<%
			End if
%>
		<%else %>
			<img src="../images/dot.gif">
		<%End if%>
		</td> -->
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
          <tr> 
            <td colspan="4" align="center" bgcolor="#FFD553"> 
        <form action= "<%=sScriptPath%>" method="Post">
        <%if TotalPage > 1 then%> 

              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../datamanager/images/dot.gif"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%
			If PageNo=1 then
				response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&sSearch=<%=Request("sSearch")%>">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&sSearch=<%=Request("sSearch")%>">Next 
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
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&sSearch=<%=Request("sSearch")%>"><%=i%></a> 
                    <%
	End if
	Next %></td>
                </tr>
              </table>
              <%End if
			  if Mid(Session("sAuthor"),22,3)="AED" then
			  %> 
			  <input type="text" name="sSearch"><input type="submit" name="search" value="Search">
			  <%End IF %>
			 </form>
            </td>
          </tr>
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
Sub chkReceive()
    Dim  Rs2,sSQL2,amount
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select sum(preceiveitem.nQuality) as Quantity from preceive,preceiveitem where preceive.spono='" & spono & "'" &_
	" and preceive.srecno=preceiveitem.srecno"
	Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if Not Rs2.EOF then
		Do while Not RS2.EOF
			amount=RS2("Quantity")
		if  amount > "0" then
			chkRec=1
		else
			chkRec=0
		end if
		Rs2.MoveNext
		Loop
	else
		chkRec=0
	End if
	Rs2.Close
	Set RS2=nothing
 End Sub

Sub chkinvoice()
    Dim  Rs2,sSQL2,amount
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select * from pInvoice where spono='" & spono & "'" 
	Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if Not Rs2.EOF then
			chkInv=1
	else
		chkInv=0
	End if
	Rs2.Close
	Set RS2=nothing
 End Sub

Sub chkReceive2()
    Dim  Rs2,sSQL2,amount
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select * from preceive where spono='" & spono & "'" 
	Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if Not Rs2.EOF then
			chkRec2=1
	else
		chkRec2=0
	End if
	Rs2.Close
	Set RS2=nothing
 End Sub
%>