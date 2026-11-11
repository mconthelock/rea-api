<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim spono,sexptype,dprdate,srequester,skey,cprtype,nstep,capprove,sexpdesc
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Prno
			spono=Request("sPono")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
		  sSQL = "SELECT * FROM  ppo,pvender WHERE ppr.sVendcode=pvender.sVendcode" &_
		" and sPono='" & spono & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
Do while Not RS.EOF
	sPono=Rs("sPono")
	dPodate=RS("dPodate")
	sRequester = RS("sRequester")
	sKey=RS("sKey")
	sEname=RS("sEname")
	sEaddress=RS("sEaddress")
	RS.MoveNext
Loop
'Session("sExptype") = sExptype
			%>
<html>
<head>
<title>Show P/O</title>
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
                  PURCHASE ORDER</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>No.</b> 
                        </font>&nbsp;<%=sprno%></td>
                      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Expense 
                        Type</b> </font>&nbsp;<%=sExpdesc%> 
                      </td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Date</b>&nbsp;</font><%response.write Show_date(dprdate)%><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                        </font></td>
                      <td width="59%">&nbsp; </td>
                    </tr>
                    <tr> 
                      <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Request 
                        By</b> </font></td>
                      <td width="25%"> <%=sRequester%> 
                      </td>
                      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Due 
                        Date</b> </font><%response.write Show_date(dprdate)%></td>
                    </tr>
                    <tr> 
                      <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Key 
                        By</b> </font>&nbsp;<%=sKey%></td>
                      <td width="59%">&nbsp;</td>
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
	Dim sScriptPath,sexptype,Color
Color="#F8D7D1"
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM ppritem  where sprno='" & sPrno & "'" &_
	"Order By nItem"

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
                <td  width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Item</font></b></font></td>
                <td width="30%" cellspacing="2" cellpadding="2" width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Name</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
                <td  width="10%" cellspacing="2" cellpadding="2" width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Price</font></b></font></td>
                <td   width="30%" cellspacing="2" cellpadding="2" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Use For</font></b></font></td>
                <td   cellspacing="2" cellpadding="2" width="8%" bgcolor="#8B2718">&nbsp;</td>
                <td   cellspacing="2" cellpadding="2" width="9%" bgcolor="#8B2718" align="right" valign="top"><img src="../images/y05.jpg" width="12" height="12"></td>
              </tr>
<% 
Item=1
Do While Not RS.EOF %>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nitem")%>&nbsp;</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sprodname")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nQuality")%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2"><%=RS("nPrice")%></td>
                <td  cellspacing="2" cellpadding="2" bgcolor="<%=Color%>" height="2"><%=RS("sUseFor")%></td>
                <td  cellspacing="2" cellpadding="2" width="8%" bgcolor="<%=Color%>" align="center" height="2" valign="top"><a href="../workflow/ppritemedt.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>&sexptype=<%=sExptype%>"><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a></td>
                <td cellspacing="2" cellpadding="2" width="9%" bgcolor="<%=Color%>" align="center" height="2"><a href="../workflow/ppritemdel.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>"><img src="../images/icondelete2.gif" width="30" height="26" alt="DELETE" border="0"> 
                </td>
              </tr>
<%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
Item=Cint(RS("nItem"))+1

 RS.MoveNext
Loop
%>
<tr>
<td colspan="6">&nbsp;</td>
<td   width="8%" bgcolor="#FFD553" align="right" height="8"><a href="../workflow/ppritemedt.asp?nItem=<%=Item%>&sPrno=<%=sPrno%>"><img src="../images/iconadd2.gif" width="30" height="26" alt="ADD" border="0"></a></td>
</tr>
            </table>
| <a href='../workflow/CreatePrApp.asp?sExptype=<%=Session("sExptype") %>&sprno=<%=sPrno%>&sRequester=<%=sRequester%>&dReqdate=<%=dprdate%>'>Create Approve</a> |
          </td>

        </tr>
 
      </table>

<%
End Sub%>