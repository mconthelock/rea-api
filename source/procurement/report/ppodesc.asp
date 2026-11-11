	<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,sRef
	Dim spono,svendcode,dpodate,srequester,skey,sRefno,sShipto,sShipby,sDelivery,sByer,sAcceptby,sPayterm,sTradeterm
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Pono,stname,sEaddress,sCurrency,sEname
spono=Request("sPono")
sSQL="Select * from ppo,pvender where ppo.svendcode=pvender.svendcode" &_
" and ppo.sPono='" & sPono & "'"
Set Rs=Server.CreateObject("ADODB.Recordset")
Rs.MaxRecords=1
Rs.Open sSQL,conn,adOpenStatic,adLockReadOnly,adCmdText
If Not Rs.EOF  then
	sPono=Rs("sPono")
	dPodate=Rs("dPodate")
	sRefno=Rs("sRefno")
	sRequester=Rs("sRequester")
	sKey=Rs("sKey")
	sRef=Rs("sRef")
	sEname=Rs("sEname")
	sEaddress=Rs("sEaddress")
	sShipto=Rs("sShipto")
	sShipby=Rs("sShipby")
	sPayterm=Rs("sPayterm")
	sTradeterm=Rs("sTradeterm")
	sDelivery=Rs("sDelivery")
	sByer=Rs("sByer")
	sCurrency=RS("sCurrency")
	sAcceptby=Rs("sAcceptby")
End if
%>
<html>
<head>
<title>PURCHASE ORDER</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
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
          <td width="2%" height="113">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="113"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td width="79%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> 
                  PURCHASE ORDER</font></b></font></td>
                <td width="21%" bgcolor="#FF9999" align="center"> 
                  <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#666666"><b><font color="#FFFFFF" size="2">DATE</font><%response.write Show_date(dpodate)%></b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"></font></div>
                </td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <table width="98%" border="0" cellspacing="2" cellpadding="1">
                    <tr> 
                      <td colspan="4" rowspan="2" bgcolor="#FCECE9"> 
                        <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Attention</b> 
                          : </font><%=sEname%></p>
                        <p>&nbsp;<%=sEaddress%></p>
                      </td>
                      <td width="29%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Purchase 
                        Order No. </b></font><br>
                        <b>
                        <%=sPono%></b></td>
                    </tr>
                    <tr> 
                      <td width="29%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Reference 
                        No. </b></font><br>
                        <b>
                        <a href="../workflow/showprdesc.asp?sprno=<%=sRefno%>"><%=sRefno%></a></b></td>
                    </tr>
                    <tr> 
                      <td colspan="3" rowspan="6" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Shipping 
                        mark </b> <br>
                        <b></b>&nbsp;</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                        </font></td>
                      <td colspan="2" height="5" bgcolor="#FCECE9"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Ship 
                        To: </b></font><font color="#000000"><%=sShipto%></font> 
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="2" bgcolor="#FCECE9"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Ship 
                        By: </b></font><font color="#000000"><%=sShipby%></font></td>
                    </tr>
                    <tr> 
                      <td width="46%" rowspan="2" bgcolor="#FCECE9"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Terms 
                        Of Payment: </b></font><font color="#000000"><%=sPayterm%></font></td>
                      <td width="29%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Currency: 
                        </b></font><font color="#000000"><%Show_Currency(sCurrency)%></font></td>
                    </tr>
                    <tr> 
                      <td width="29%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery:</b></font> 
                        <font color="#000000"><%=sDelivery%></font></td>
                    </tr>
                    <tr> 
                      <td width="46%" rowspan="2" bgcolor="#FCECE9"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Terms 
                        Of Trade: </b></font><font color="#000000"><%=sTradeterm%></font></td>
                      <td width="29%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Buyer 
                        Code: </b></font><font color="#000000"><%=sbyer%></font></td>
                    </tr>
                    <tr> 
                      <td width="29%">&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2"><% Show_Item()%></td>
              </tr>
            </table>
            <p><br>
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="workflow.asp">WorkFlow Manager</a></font> <font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
              </font></p>
          </td>
          <td width="2%" height="113">&nbsp;</td>
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
	Dim sScriptPath,sexptype,Color
Color="#F8D7D1"
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM ppoitem  where spono='" & sPono &_
	"' Order By sProdcode"
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
'response.write RS2("MaxItem")	
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="100%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">

              <tr bgcolor="#DB4B37"> 
                <td  width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Code</font></b></font></td>
                <td width="35%" cellspacing="2" cellpadding="2" width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Name</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
                <td  width="5%" cellspacing="2" cellpadding="2" width="5%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">U/M</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">UPrice</font></b></font></td>
                <td   width="10%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Amount</font></b></font></td>

                <td   cellspacing="2" cellpadding="2" width="9%" bgcolor="#8B2718" align="right" valign="top"><img src="../images/y05.jpg" width="12" height="12"></td>
              </tr>
<% 
Item=1
Do While Not RS.EOF %>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdcode")%>&nbsp;</td>
                <td  bgcolor="<%=Color%>" height="2"><A HREF="JavaScript:OpenNewWindow('../workflow/ppoitemdesc.asp?spono=<%=RS("spono")%>&sprodcode=<%=Rs("sProdcode")%>','SearchProduct','width=500,height=500')"><%=RS("sprodname")%></a></td>
                <td  bgcolor="<%=Color%>" height="2" align="right"><%=FormatNumber(RS("nQuality"))%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2" align="right"><%=FormatNumber(RS("nPrice"), 4)%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2" align="right"><%=FormatNumber(RS("nAmount"), 4)%></td>
                <td  cellspacing="2" cellpadding="2" width="8%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
                </td>
              </tr>
<%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
'Item=Cint(RS("nItem"))+1
 RS.MoveNext
Loop
Rs.Close
	sSQL = "SELECT sCurrency,Sum(nAmount) sumamount FROM ppoitem,pCurrency  where sPono='" & sPono & "'" &_
	" and ppoitem.sCurcode=pCurrency.sCurcode" &_
	" Group By sCurrency" &_
	" Order By sCurrency"
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<tr>
<td colspan="9" align="center">
<%	IF Not Rs.EOF then %>
	<table width="50%">
		<tr><td colspan="2" align="center" bgcolor="teal"><Font color="#FFFFFF"><B>Total</B></Font></td></tr>
		<tr><td  width="50%" align="right" bgcolor="#F1B3A7"><Font color="#1111FF"><b><%=Formatnumber(RS("sumamount"), 4)%></b></font></td><td bgcolor="#F1B3A7"><%=RS("sCurrency")%></td></tr>
	</table>
<%	End if%>
</td>
</tr>
            </table>

          </td>

        </tr>
 
      </table>

<%
End Sub%>