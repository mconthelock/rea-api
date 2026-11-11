<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprno,sexptype,dprdate,srequester,skey,cprtype,nstep,capprove,sexpdesc,Amount
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Prno,sCashreq,sDelivery
	Dim dDelivery
	sprno=Request("sPrno")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
''		  sSQL = "SELECT * FROM  ppr,pexpenstype WHERE ppr.sexptype=pexpenstype.sexptype" &_
''			" and sprno='" & sprno &"'"
'modify by Kallayanee On 26/09/2008
'date return in thai year change sql select  to to_char(dprdate)
		sSQL = "SELECT sPrno,sExpdesc,ppr.sExptype, to_char(dPrdate,'dd/mon/yyyy') dPrdate,sDelivery, dDelivery, sRequester, sKey, cPrtype, sCashReq " _
		 & " FROM  ppr,pexpenstype WHERE ppr.sexptype=pexpenstype.sexptype " &_
		" and sprno='" & sprno &"'"
'
'			response.write sSQL
'			response.end
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
'	  response.end
Do while Not RS.EOF
	sPrno=Rs("sPrno")
	sExpdesc=RS("sExpdesc")
	sExptype = RS("sExptype")
	dPrdate=RS("dPrdate")
	sDelivery=Rs("sDelivery")
	dDelivery = rs("dDelivery")
	sRequester=RS("sRequester")
	sKey=RS("sKey")
	cPrtype=RS("cPrtype")
	sCashReq=Rs("sCashReq")
RS.MoveNext
Loop
Session("sExptype") = sExptype

			%>
<html>
<head>
<title>Show P/R</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</Script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- Start table 1 -->
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2">&nbsp;</td>
    <td width="63%" height="2" align="center" valign="top"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td width="36%" height="2" align="center"><img src="../security/images/spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td width="1%" height="2"><img src="../security/images/spacer.gif" width="1" height="1"></td>
    <td colspan="2" height="2" align="center" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font> 
<!-- start table 2 -->
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">...</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top" > 
<!-- start table 3 -->
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  PURCHASE REQUISITION</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                    
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>No.</b> 
                        </font>&nbsp;<%=sprno%></td>
                      <td width="51%" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Expense 
                        Type</b> </font>&nbsp;<%=sExpdesc%> </td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Date&nbsp;</b></font><%response.write Show_date(dprdate)%><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                        </font></td>
                      <td width="51%" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery Date :</b> </font><%=sDelivery%><%if dDelivery <> "" then response.write "<br><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2"" color=""#CC3333""><b>Expected date : </b></font>" & Show_date(dDelivery) end if%>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Request 
                        By </b></font>(<%=sRequester%>) <%Show_Empno(sRequester)%> <%Show_Empdesc()%></td>
                      
                      <td width="51%" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Cash 
                        Requirement No. </b></font><%Response.write sCashReq%></td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Key 
                        By</b> </font>&nbsp;(<%=sKey%>)<%Show_Empno(sKey)%></td>
                      <td width="51%" colspan="2"><%if cPrtype="1" then %>
						<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>
						<Blink>Urgent Requisition </Blink></b></font><%End if%></td>
                    </tr>
                    <tr> 
                      <td width="15%"><font color="#CC3333"><b>Quotation No.</b></font></td>
                      <td colspan="3"><%Show_file()%>&nbsp;</td>
					  <td align="right"><A HREF="ppritemdescedt.asp?sPrno=<%=sPrno%>">Add Quotation</a>
						<%Session("sprno")=sprno%>
						</td>
                    </tr>
                  </table>
				<!-- End table 3 -->
                </td> 
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;<% Show_Item()%></td>
              </tr>
            </table>
  					<br>
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../workflow/workflow.asp">Workflow Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font>
          </td>
          <td width="2%" >&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
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
</table> <!-- end table 1 -->
<br>
<br>
</body>
</html>
<%
	Sub Show_Item()
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,Color,count
Color="#F8D7D1"
count=1
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

'		sSQL = "delete from ppritem" _
'		& " where exists ( select svendcode from  pvender " _ 
'		& " where  pvender.svendcode = ppritem.svendcode " _
'		& " and    pvender.cstatus = '0') and ppritem.sprno = '" & sPrno & "'"
'		response.write sSQL
'		response.end
'		Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		
	sSQL = "SELECT  * FROM ppritem  where sprno='" & sPrno & "'" &_
	"Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
'response.write RS2("MaxItem")	
Item=1
%>
<!-- start table 4 -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td width="100%" align="left" valign="top" bgcolor="#FFD553"> 
<!-- start table 5 -->
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr align="right"> 
                <td colspan="10"><a href="searchprcopy.asp?sNewPrno=<%=sPrno%>&sExptype=<%=sExptype%>">Search P/R</a> </td>
              </tr>

              <tr bgcolor="#DB4B37"> 
                <td  width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Item</font></b></font></td>
                <td width="25%" cellspacing="2" cellpadding="2" width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Name</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718" align="right" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
                <td  width="5%" cellspacing="2" cellpadding="2" width="5%"  bgcolor="#8B2718" align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">U/M</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%" bgcolor="#8B2718" align="right" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">UPrice</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%" bgcolor="#8B2718" align="right" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Amount</font></b></font></td>
                <td  width="3%" cellspacing="2" cellpadding="2" width="3%" bgcolor="#8B2718" align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Cur</font></b></font></td>
<!--                <td   width="10%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Use For Sec.</font></b></font></td> -->
                <td  colspan="1" width="10%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Comment</font></b></font></td>
                <td   cellspacing="2" cellpadding="2" width="8%" bgcolor="#8B2718">&nbsp;</td>
                <td   cellspacing="2" cellpadding="2" width="9%" bgcolor="#8B2718" align="right" valign="top"><img src="../images/y05.jpg" width="12" height="12"></td>
              </tr>
<% 


Do While Not RS.EOF %>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nItem")%>&nbsp;</td>
                <td  bgcolor="<%=Color%>" height="2"><A HREF="JavaScript:OpenNewWindow('../workflow/ppritemdesc.asp?sprno=<%=RS("sprno")%>&nitem=<%=Rs("nItem")%>&sExptype=<%=sExptype%>','ShowPrdesc','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')"><%=RS("sprodname")%></a></td>
                <td  bgcolor="<%=Color%>" height="2" align="right"><%=FormatNumber(RS("nQuality"),3)%></td>
                <td  bgcolor="<%=Color%>" height="2" align="left"><%=RS("sUnitcode")%></td>
                <td  cellspacing="2" align="right"  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 3)%></td>
                <td  cellspacing="2" align="right" bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nAmount"), 3)%></td>
                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" width="1%"><img src="../images/dot.gif"><%=RS("sCurcode")%></td>
<!--                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2"><%=RS("sUseFor")%><img src="../images/dot.gif"></td>
                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" width="10%"><img src="../images/dot.gif"><A HREF="JavaScript:OpenNewWindow('../workflow/reject.asp?sPrno=<%=sPrno%>&nItem=<%=RS("nItem")%>&sExptype=<%=Request("sExptype")%>','SearchProduct','width=500,height=500')"><%if RS("sComment")<>"" then %>View<%End if%></a></td>-->
                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" width="10%"><img src="../images/dot.gif"><%=RS("sComment")%></td>
                <td  cellspacing="2" cellpadding="2" width="8%" bgcolor="<%=Color%>" align="center" height="2"><a href="../workflow/ppritemedt2.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>&sexptype=<%=sExptype%>"><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a></td>
                <td cellspacing="2" cellpadding="2" width="9%" bgcolor="<%=Color%>" align="center" height="2"><a href="../workflow/ppritemdel.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>"><img src="../images/icondelete2.gif" width="30" height="26" alt="DELETE" border="0"></a>
                </td>
              </tr>
<%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
count=count+1
Item=Cint(RS("nItem"))+1
Amount=Amount+(cDbl(RS("nQuality"))*cDbl(RS("nPrice")))
 RS.MoveNext
Loop
Rs.Close
		
	sSQL = "SELECT sCurcode,Sum(nAmount) sumamount FROM ppritem  where ppritem.sprno='" & sPrno & "'" &_
	" and cstatus<>'0'" &_
	" Group By sCurcode" &_
	" Order By sCurcode"
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<tr>
<td colspan="9" align="Right">&nbsp;</td>

<td   width="8%" bgcolor="#FFD553" align="center" height="8"><a href="../workflow/ppritemedt2.asp?nItem=<%=Item%>&sPrno=<%=sPrno%>&sExptype=<%=Session("sExptype") %>"><img src="../images/iconadd2.gif" width="26" height="26" alt="ADD" border="0"></a></td>

</tr>
<tr>
<td colspan="10">
<%IF Not RS.EOF then%>
<center>

<%	Do while Not RS.EOF
%>
<!-- start table 6 -->
<table width="50%" border="0">
	<tr>
		<td colspan="2" align="center" bgcolor="teal"><Font color="#FFFFFF"><B>Total</b></Font></td>
	</tr>
	<tr>
		<td  align="Right" bgcolor="#F1B3A7" colspan="1" width="50%">&nbsp;<Font color="#1111FF"><B><%=FormatNumber(RS("sumAmount"), 3)%></B></font></td>
		<td  align="Center" bgcolor="#F1B3A7">&nbsp;<%show_Currency(RS("sCurcode"))%></td>
	</tr>
<%
		RS.MoveNext
		Loop
%>
      </table>
<!-- stop table 6 -->
	<%End if%>

</center></td>
</tr>
</table>
<!-- stop table 5 -->
</tr>
</table>
<!-- stop table 4 -->
<br>
<%IF Not RS.EOF OR Item > 1 then%>
| <a href='../workflow/CreatePrApp.asp?sExptype=<%=Session("sExptype") %>&sprno=<%=sPrno%>&sRequester=<%=sRequester%>&dReqdate=<%=dprdate%>&cPrtype=<%=cPrtype%>'>Send To Be Approved</a> 
| <a href='../workflow/CreatePrApp.asp?sExptype=<%=Session("sExptype") %>&sprno=<%=sPrno%>&sRequester=<%=sRequester%>&dReqdate=<%=dprdate%>&cPrtype=<%=cPrtype%>&go=next'>Send To Be Approved and Add Next P/R</a>
<%End if%>| <a href="../workflow/pprdel.asp?sprno=<%=sPrno%>">Delete P/R</a>|
     &nbsp;     
<!--
			</td>
        </tr>
      </table>
-->
<%
End Sub
Sub Show_Emp(sEmpno)
    Dim  Rs,sSQL,sStatus,sScriptPath,smode
	Dim scurcode,scurrency
	Dim img_sEmpno, img_sName, img_sSurname, img_sTName,  img_sTSurname, img_sLogin ,img_sPassword
	Dim img_sDivision, img_sDepartment, img_sSection, img_sPosition
	Dim sDivision, sDepartment, sSection, sPosition
	Dim sDivcode, sDepcode, sSeccode, sPoscode
	Dim sDivc, sDepc, sSecc, sPO, sName, sSurname, sTName, sTSurname
	Dim sLogin, sPassword, sPassword1
	Dim bError, sbread
	Dim sErrImage
	bError = False
	sErrImage = "<IMG SRC='../images/icerror.gif'>"
sScriptPath = Request.ServerVariables("SCRIPT_NAME")


  sSQL = "SELECT * FROM vaEmployee WHERE sEmpno ='" & sEmpno & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sEmpno = Rs("sEmpno")
		sDivcode = Rs("sDivcode")
		sDivision=RS("sDivision")
		sDepcode = Rs("sDepcode")
		sDepartment=RS("sDepartment")
		sSeccode = Rs("sSeccode")
		sSection=Rs("sSection")
		sPoscode = Rs("sPoscode")
		sPosition=RS("sPosition")
		sName = Rs("sName")
		sSurname = Rs("sSurname")
		sTName = Rs("sTName")
		sTSurname = Rs("sTSurname")
		sLogin = Rs("sLogin")
		sPassword = RS("sPassword")
Response.write sName
	End If
End Sub
%>
<%
Sub show_file()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from ppritemdesc Where sprno='" & sprno & "'"

		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
%><a href="../workflow/ppritemdescimages/<%=RS2("sfile")%>" target="_Blank">
<%					Response.Write(Rs2("sfile") & "</a>&nbsp;&nbsp;&nbsp;")
					Response.write(RS2("sdesc") & "&nbsp;&nbsp;&nbsp;&nbsp;")
%><a Href="delfile.asp?sprno=<%=sprno%>&nitem=<%=RS2("nitem")%>&sExptype=<%=Request("sExptype")%>&sitemno=<%=RS2("sitemno")%>">Del</a><br>
<%
			Rs2.MoveNext
		Loop
	End IF
End Sub

Sub Show_Empdesc()
	Dim RS2,sSQL2

	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select *  from vAemployee where  sEmpno='" & sRequester & "'"

	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF  Not Rs2.EOF then
			response.write "<font size=1> " & RS2("sSeccode") & "/"
			response.write RS2("sDepcode") & "</font>"
		End if

End Sub
%>
