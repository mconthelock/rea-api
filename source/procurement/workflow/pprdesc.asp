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
	Dim Item,Prno,sCashreq
			sprno=Request("sPrno")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
		  sSQL = "SELECT * FROM  ppr,pexpenstype WHERE ppr.sexptype=pexpenstype.sexptype" &_
			" and sprno='" & sprno &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
Do while Not RS.EOF
	sPrno=Rs("sPrno")
	sExpdesc=RS("sExpdesc")
	sExptype = RS("sExptype")
	dPrdate=RS("dPrdate")
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
                      <td width="51%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b></b> </font>
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Request 
                        By </b></font>(<%=sRequester%>) <%Show_Empno(sRequester)%></td>
                      
                      <td width="51%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Cash 
                        Requirement No. </b></font><%Response.write sCashReq%></td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Key 
                        By</b> </font>&nbsp;(<%=sKey%>) <%Show_Empno(sKey)%></td>
                      <td width="51%"><%If cprtype="1" then%>
					<Blink><font color="#CC3333"><b>Urgent Requisition </b></font></Blink><%End if%></td>
                    </tr>
                    <tr> 
                      <td colspan="4" align="center"></td>
                    </tr>
                  </table>
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
</table>
<br>
<br>
</body>
</html>
<%
	Sub Show_Item()
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,Color
Color="#F8D7D1"
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  ppritem.*, punit.SUNITdesc,PPRODUCT.HAZARDSTATUS  FROM ppritem LEFT JOIN punit ON punit.sunitcode = ppritem.sunitcode LEFT JOIN PPRODUCT ON PPRODUCT.Sprodcode = ppritem.Sprodcode where ppritem.sprno='" & sPrno & "'" &_
	" and punit.cstatus = '1' and PPRODUCT.SEXPTYPE='"&sexptype&"'" &_
	"Order By nItem"
	'response.write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
'response.write RS2("MaxItem")	
Item=1
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td width="100%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">

              <tr bgcolor="#DB4B37"> 
                <td  width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Item</font></b></font></td>
                <td width="30%" cellspacing="2" cellpadding="2" width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Name</font></b></font></td>
				<td  width="5%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718" align="left" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Hazard Substance</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718" align="right" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
				 <td  width="10%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718" align="left" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">U/M</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%" bgcolor="#8B2718" align="right" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">UPrice</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%" bgcolor="#8B2718" align="right" ><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Amount</font></b></font></td>
                <td  colspan="1" width="5%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Cur</font></b></font></td>
                <td   width="20%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Comment</font></b></font></td>
                <td   cellspacing="2" cellpadding="2" width="9%" bgcolor="#8B2718" align="right" valign="top"><img src="../images/y05.jpg" width="12" height="12"></td>
              </tr>
<% 


Do While Not RS.EOF %>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nitem")%>&nbsp;</td>
                <td  bgcolor="<%=Color%>" height="2"><A HREF="JavaScript:OpenNewWindow('../workflow/ppritemdesc.asp?sprno=<%=sprno%>&nitem=<%=Rs("nItem")%>','SearchProduct','')"><%=RS("sprodname")%></a></td>
				<td  bgcolor="<%=Color%>" height="2" align="right"><%%></td>
                <td  bgcolor="<%=Color%>" height="2" align="right"><%=FormatNumber(RS("nQuality"),3)%></td>
                <td  bgcolor="<%=Color%>" height="2" align="left"><%=RS("sUnitdesc")%></td>
                <td  cellspacing="2" align="right"  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 4)%></td>
                <td  cellspacing="2" align="right" bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nAmount"),4)%></td>
                <td  colspan="1" cellpadding="2" bgcolor="<%=Color%>" height="2" width="10%"><img src="../images/dot.gif"><%=RS("sCurcode")%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2"><%=RS("sComment")%><img src="../images/dot.gif"></td>
                <td  cellspacing="2" cellpadding="2" width="9%" bgcolor="<%=Color%>" align="center" height="2"></td>
              </tr>
<%
	if Color="#F8D7D1" then
		Color="#F1B3A7"
	else
		Color="#F8D7D1"
	end if
		Item=Cint(RS("nItem"))+1
		Amount=Amount+(cDbl(RS("nQuality"))*cDbl(RS("nPrice")))
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
<td colspan="9"  >&nbsp;</td></tr>
</table>
<center>
<table width="50%">
<tr><td colspan="2" align="center" bgcolor="teal"><Font color="#FFFFFF"><B>Total</b></Font></td>
<%IF Not RS.EOF then
	Do while Not RS.EOF
%>
<tr>
<td  align="Right" bgcolor="#F1B3A7" colspan="1" width="40%">&nbsp;<Font color="#1111FF"><B><%=FormatNumber(RS("sumAmount"), 4)%></B></font></td>
<td  align="Center" bgcolor="#F1B3A7">&nbsp;<%show_Currency(RS("sCurcode"))%></td>
</tr>
<%
	RS.MoveNext
	Loop
else%>
<td colspan="3"  >&nbsp;</td>
<%End if%>
<td colspan="2" align="Right">&nbsp;</td>

<td   width="8%" bgcolor="#FFD553" align="center" height="8"></td>
</tr> 
            </table>
     &nbsp;     
			</td>
        </tr>
      </table>

<%
End Sub%>