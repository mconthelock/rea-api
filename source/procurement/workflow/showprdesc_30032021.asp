<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprno,sexptype,dprdate,srequester,skey,cprtype,nstep,capprove,sexpdesc,sEmpno
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,cApprov,sCashReq,chk1
	Dim Item,Prno,sStepdesc,Amount,chk2,chk3,sDelivery
	Dim dDelivery
			sprno=Request("sPrno")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
		  sSQL = "SELECT * FROM  ppr,pexpenstype WHERE ppr.sexptype=pexpenstype.sexptype" &_
			" and sprno='" & sprno &"'"
			
	'modify by Kallayanee On 26/09/2008
	'date return in thai year change sql select  to to_char(dprdate)
	sSQL = "SELECT sPrno,sExpdesc,ppr.sExptype, to_char(dPrdate,'dd/mm/yyyy') dPrdate,sDelivery, dDelivery, sRequester, sKey, cPrtype, sCashReq " _
	 & " FROM  ppr,pexpenstype WHERE ppr.sexptype=pexpenstype.sexptype " &_
	" and sprno='" & sprno &"'"
	Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	
Do while Not RS.EOF
	sPrno=Rs("sPrno")
	sExpdesc=RS("sExpdesc")
	sExptype = RS("sExptype")
	dPrdate=RS("dPrdate")
	sRequester=RS("sRequester")
	sDelivery=RS("sDelivery")
	dDelivery = rs("dDelivery")
	sKey=RS("sKey")
	cprtype=RS("cprtype")
	sCashReq=Rs("sCashReq")
RS.MoveNext
Loop
Rs.Close
Session("sExptype") = sExptype
			%>
<html>
<head>
<title>ExpensStep</title>
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
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">...</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top"><b><img src="../images/plus.gif" width="9" height="9"> 
            WORK FLOW MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top" > 
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
                      <td width="51%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Expense
                        Type</b> </font>&nbsp;<%=sExpdesc%> </td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Date&nbsp;</b></font><%response.write Show_date(dprdate)%><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                        </font></td>
                      <td width="51%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery Date :</b> </font><%=sDelivery%><%if dDelivery <> "" then response.write "<br><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2"" color=""#CC3333""><b>Expected date : </b></font>" & Show_date(dDelivery) end if%>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Request 
                        By </b></font>(<%=sRequester%>)
						<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sRequester %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
						<%Show_Empno(sRequester)%></a><%Show_Empdesc%> </td>
                      <td width="51%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Cash 
                        Requirement No. </b></font><%Response.write sCashReq%></td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Key 
                        By</b> </font>&nbsp;(<%=sKey%>)
						<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=skey %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
						<%Show_Empno(sKey)%></a></td>
                      <td width="51%"><%if cPrtype="1" then %>
						<font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>
						<Blink>Urgent Requisition </Blink></b></font><%End if%></td>
                    </tr>
<%
					chk()
					chkprapp()
					chkkeyreq()
%>
                    <tr> 
                      <td width="15%"><font color="#CC3333"><b>Quotation.</b></font></td>
                      <td colspan="2"><%Show_file()%>&nbsp;</td>
					  <td align="right"><img src="../images/dot.gif">
					<%If chk3=1 then%>
						<A HREF="ppritemdescedt.asp?sPrno=<%=sPrno%>">Add Quotation</a>
						<%Session("sprno")=sprno%>
					<%End if%>
						</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;
				<% 
					Show_Item(sExptype)%>
				</td>
              </tr>
            </table>
            <br>
             <%
					if chk3<>1 then %>
            <table border="0" width="100%" cellspacing="1" cellpadding="0">
              <tr bgcolor="#8B2718"> 
                <td width="3%" height="16" >&nbsp;</td>
                <td width="27%" height="16"><b><font color="Red">Step No.</font></b></td>
                <td width="39%" height="16"><font color="Red"><b>Approve By</b></font></td>
                <td width="18%" height="16" align="center"><font color="Red"><b>Date</b></font></td>
                <td width="13%" height="16" align="center"><font color="Red"><b>Time</b></font></td>
              </tr>
 
				<%Showdesc()%>			
            </table>
				<%	End if	%> 
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
</table>
<br>
<br>
</body>
</html>
<%
	Sub Show_Item(sexptype)
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,Color,count
Color="#F8D7D1"
count=1
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	'sSQL = "SELECT  * FROM ppritem  where sprno='" & sPrno & "'" &_
	sSQL = "SELECT  ppritem.*, punit.SUNITDESC,PPRODUCT.SEPRODNAME,PPRODUCT.HAZARDSTATUS FROM ppritem LEFT JOIN punit ON punit.sunitcode = ppritem.sunitcode " &_
	"LEFT JOIN PPRODUCT ON PPRODUCT.Sprodcode = ppritem.Sprodcode  " &_
	"where ppritem.sprno ='" & sPrno & "' and PPRODUCT.SEXPTYPE = '" & sExptype &"'Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
'response.write RS2("MaxItem")	
'response.write sSQL
chkkeyreq()
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="100%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
             <tr align="right"> 
                <td colspan="10">
			<%If chk3="1" then%>
				<a href="searchprcopy.asp?sNewPrno=<%=sPrno%>&sExptype=<%=sExptype%>">Search P/R</a> 
			<%End IF%>
				</td>
              </tr>
              <tr bgcolor="#DB4B37"> 
                <td  width="3%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Item</font></b></font></td>
                <td width="20%" cellspacing="2" cellpadding="2"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Name</font></b></font></td>
				<td  width="3%" cellspacing="2" cellpadding="2" width="3%"  bgcolor="#8B2718" align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Hazard Substance</font></b></font></td>
                <td  width="6%" cellspacing="2" cellpadding="2"   bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
                <td  width="3%" cellspacing="2" cellpadding="2" width="3%"  bgcolor="#8B2718" align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">U/M</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">UPrice</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Amount</font></b></font></td>
                <td  colspan="1" width="3%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Cur</font></b></font></td>
                <td width="3%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="1">Received Budget(ß.)</font></b></font>
					</td>
                <td width="3%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="1">Used Budget(ß.)</font></b></font>
					</td>
<!--                <td  colspan="1" width="10%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Use For Sec.</font></b></font></td>-->
                <td  colspan="1" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Comment</font></b></font></td>
                <td  colspan="2" bgcolor="#8B2718" align="right" valign="top"><img src="../images/y05.jpg" width="12" height="12"></td>
              </tr>
<% 
Dim budget
Item=1
Do While Not RS.EOF 
budget = split(Show_Budget(RS("sProdCode"), dPrdate, sExptype),",")
%>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2" width="5%"><%=RS("nItem")%>&nbsp;</td>
                <td  bgcolor="<%=Color%>" height="2" width="30%"><A HREF="JavaScript:OpenNewWindow('../workflow/ppritemdesc.asp?sprno=<%=RS("sprno")%>&nitem=<%=Rs("nItem")%>&sExptype=<%=sExptype%>','ShowPrdesc','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')"><%=RS("sprodname")%></a></td>
				<td  bgcolor="<%=Color%>" height="2" align="center" width="6%">
				<%if rs("HAZARDSTATUS")="1" then %>
				YES
				<%elseif rs("HAZARDSTATUS")="0" then %>
				No
				<%else%>
				
				<%end if%>
				</td>
                <td  bgcolor="<%=Color%>" height="2" align="right" width="6%"><%=FormatNumber(RS("nQuality"), 3)%></td>
                <td  bgcolor="<%=Color%>" height="2" align="left"><%=RS("sUnitdesc")%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2" align="right" width="10%"><%=FormatNumber(RS("nPrice"), 4)%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2" align="right" width="10%"><%=FormatNumber(RS("nAmount"), 4)%></td>
                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" width="3%"><img src="../images/dot.gif"><%=RS("sCurcode")%></td>
                <td  bgcolor="<%=Color%>" height="2" valign="top"><%=FormatNumber(budget(0),0)%></td>				
                <td  bgcolor="<%=Color%>" height="2" valign="top">
				<%	
					if cCur(budget(0)) < cCur(budget(1)) then
						response.write "<font color='#FF0000'>"
					else
						response.write "<font color='#000000'>"		
					end if
					response.write FormatNumber(budget(1),0)
					response.write "<font>"
				%>
				</td>								
				
<!--                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" width="10%"><img src="../images/dot.gif"><%=RS("sUseFor")%></td>
                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" width="10%" align="center"><img src="../images/dot.gif"><A HREF="JavaScript:OpenNewWindow('../workflow/reject.asp?sPrno=<%=sPrno%>&nItem=<%=RS("nItem")%>&sExptype=<%=Request("sExptype")%>','SearchProduct','width=500,height=500')"><%if RS("sComment")<>"" then %>View<%End if%></a></td>-->
                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" align="left"><img src="../images/dot.gif"><%=RS("sComment")%></td>
                <td colspan="2" bgcolor="<%=Color%>" align="center" height="2" valign="Center"<%if chk1="1" then %> width="19%" <%End if%>>
				<% If chk1="1" then %>
					<a href="ppritemedt2.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>&sExptype=<%=sExptype%>"><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a>
					<a href="ppritemdel.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>&sExptype=<%=sExptype%>"><img src="../images/icondelete2.gif" width="30" height="26" alt="DELETE" border="0"></a>
				<% Else %>
				<img src="../images/dot.gif">
					<% 	if cint(RS("cstatus"))= 0 then%>
					<a href="reject.asp?sprno=<%=RS("sprno")%>&nitem=<%=RS("nitem")%>" target="_Blank"><%=RS("sUsefor")%></a>
				<%		End if
					End if%>
				</td>
                </td>
              </tr>
<%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
Count=count+1
Item=Cint(RS("nItem"))+1
Amount=Amount+((Cdbl(RS("nQuality")))*(Cdbl(RS("nPrice"))))
 RS.MoveNext
Loop
	RS.Close
	sSQL = "SELECT  sCurcode,sum(namount) as sumamount FROM ppritem  where sprno='" & sPrno & "'" &_
	" and cstatus <>'0'" &_
	" Group By sCurcode"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<tr>
<td colspan="9"  >&nbsp;</td>
<%If chk1="1" then%>
<td   width="8%" bgcolor="#FFD553" align="center" height="8"><a href="../workflow/ppritemedt2.asp?nItem=<%=Item%>&sPrno=<%=sPrno%>&sExptype=<%=Session("sExptype") %>"><img src="../images/iconadd2.gif" width="26" height="26" alt="ADD" border="0"></a></td>
<%End if%>
</tr>
</table>
<center>
<table width="50%">

<%IF Not RS.EOF then%>
<tr><td colspan="2" align="center" bgcolor="teal">
<Font color="#FFFFFF"><B>Total</b></Font></td>
<%	Do while Not RS.EOF
%>
<tr>
<td  align="Right" bgcolor="#F1B3A7" colspan="1" width="50%">&nbsp;<Font color="#1111FF"><B><%=FormatNumber(RS("sumAmount"), 4)%></B></font></td>
<td  align="Center" bgcolor="#F1B3A7">&nbsp;<%show_Currency(RS("sCurcode"))%></td>
</tr>
<%
	RS.MoveNext
	Loop
else%>
<td colspan="3"  >&nbsp;</td>
<%End if%>
</tr> 
</table>
</center>
<%If chk3="1" then%>
<%IF Not RS.EOF OR Item > 1 then%>
| <a href='../workflow/CreatePrApp.asp?sExptype=<%=Session("sExptype") %>&sprno=<%=sPrno%>&sRequester=<%=sRequester%>&dReqdate=<%=dprdate%>&cPrtype=<%=cPrtype%>'>Send To Be Approved</a>
| <a href='../workflow/CreatePrApp.asp?sExptype=<%=Session("sExptype") %>&sprno=<%=sPrno%>&sRequester=<%=sRequester%>&dReqdate=<%=dprdate%>&cPrtype=<%=cPrtype%>&go=next'>Send To Be Approved and Add Next P/R</a>
 <%End if%>| <a href="../workflow/pprdel.asp?sprno=<%=sPrno%>">Delete P/R</a>|

     &nbsp;    
<%End if%> 
          </td>

        </tr>
 
      </table>

<%
End Sub
Sub Showdesc()
Dim RS2,sSQL2,RS3
Dim sEmpno,dappdate,tapptime,nPrStep,sEmpnoapp,sComment
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select nStep from ppr where   sprno='" & sprno &"'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
	Do While Not Rs2.EOF
		nPrStep=Rs2("nStep")
	Rs2.MoveNext
	Loop
	Rs2.Close
'	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
		  sSQL2 = "SELECT * FROM  pprapp,pFlowstep WHERE  pprapp.sstep=pFlowstep.sstep and sprno='" & sprno &"'" &_
			" Order By nstep"
'			response.write sSQL2
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
Do while Not RS2.EOF
	sEmpno=Rs2("sEmpno")
	sEmpnoapp=RS2("sEmpnoapp")
	nstep=RS2("nstep")
	sStepdesc=RS2("sStepdesc")
	dappdate=RS2("dappdate")
	tapptime=RS2("tapptime")
	cApprov=Rs2("cApprov")
	sComment=RS2("sComment")
if nstep<>"" and nprstep<>"" then
If Cint(nStep)=Cint(nPrStep) and capprov<>"1" then
%>
<tr><td><img src='../images/istepappi.gif'></td><td><%=sStepdesc%></td>
<td><Font Color=red>(<%=sEmpno%>)<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpno %>','SearchProduct','width=500','height=250')"><%Show_Empno(sEmpno)%></A><%if sEmpnoapp<>"" then%> /(<%=sEmpnoapp%>)<%End if%>
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=RS2("sEmpnoapp") %>','showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpnoapp)%></A>
</font></td><td align="Center"><%=dappdate%></td><td align="Center"><%= tapptime%></td></tr>
<%
Elseif cApprov="1" or cApprov="5" then %>
<tr><td><img src='../images/istepapp.Gif'></td><td><%=sStepdesc%></td><td><Font Color=red>
<%if sEmpNo <> sEmpNoApp and not isNull(sEmpNoApp) then%>
<%if sEmpNo = sComment then response.write "->" end if%>
(<%=sEmpNo%>)
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpNo %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpNo)%></A>/
<%end if%>

<%if sEmpNoApp = sComment then response.write "->" end if%>
<%if sEmpNo <> sEmpNoApp and not isNull(sEmpNoApp) then%>
(<%=sEmpNoApp%>)
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpNoApp%>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpNoApp)%></A>
<%else%>
(<%=sComment%>)
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sComment%>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sComment)%></A>
<%end if%>
</Blink>

</font></td><td align="Center"><%=dappdate%></td><td align="Center"><%= tapptime%></td></tr>
<%
Else %>
<tr><td><img src='../images/istepappw.Gif'></td><td><%=sStepdesc%></td><td><Font Color=red>(<%=sEmpno%>)<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpno %>','Showempno','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%Show_Empno(sEmpno)%></A></Blink><%If sEmpnoapp <> "" then %> / (<%=sEmpnoapp%>)<%End if%>
<A HREF="JavaScript:OpenNewWindow('../Datamanager/vEmployeedesc.asp?sEmpno=<%=sEmpnoapp %>','SearchProduct','width=500','height=250')">
<%Show_Empno(sEmpnoapp)%></A>
</font></td><td align="Center"><%=dappdate%></td><td align="Center"><%= tapptime%></td></tr>
<%
End if
End if
RS2.MoveNext
Loop
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

Sub chk()
	Dim RS2,sSQL2,RS3
	chk1=0
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select *  from ppr where  sprno='" & sprno & "'" &_
	" and cApprove='9'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF Not Rs2.EOF then
			chk1="1"
		End if

End Sub


Sub chkkeyreq()
	Dim RS2,sSQL2,RS3
	chk2=0
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select *  from ppr where  sprno='" & sprno & "'" & " and (sKey='" & Session("sEmpno") &"'" &_
" OR sRequester='" & Session("sEmpno") & "')"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF Rs2.EOF then
			chk2="1"
		End if

End Sub

Sub chkprapp()
	Dim RS2,sSQL2,RS3
	chk3=0
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select *  from pprapp where  sprno='" & sprno & "'"
'	response.write sSQL2
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF  Rs2.EOF then
			chk3="1"
		End if

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
%>

<a href="JavaScript:OpenNewWindow('../workflow/ppritemdescimages/<%=Rs2("sfile")%>','Showimage','toolbar= 0,location= 0,directories= 0,status= 0,menubar= 0,scrollbars= 1,resizable= 1,copyhistory= 0,width=600,height=500')">
<%					Response.Write(Rs2("sfile") & "</a>&nbsp;&nbsp;&nbsp;")
					Response.write(RS2("sdesc") & "&nbsp;&nbsp;&nbsp;&nbsp;")
		if chk3=1 then %>
			<a Href="delfile.asp?sprno=<%=sprno%>&nitem=<%=RS2("nitem")%>&sExptype=<%=Request("sExptype")%>&sitemno=<%=RS2("sitemno")%>">Del</a>
<%		End if
		Response.write "<br>"
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