	<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<HTML>
<HEAD>
<TITLE>PO1</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-874">


<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim spono,svendcode,dpodate,srequester,skey,sRefno,sShipto,sShipby,sDelivery,sByer,sAcceptby,sPayterm,sTradeterm
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Pono,stname,sEaddress,sCurrency
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
	sTname=Rs("sTname")
	sEaddress=Rs("sEaddress")
	sShipto=Rs("sShipto")
	sShipby=Rs("sShipby")
	sPayterm=Rs("sPayterm")
	sTradeterm=Rs("sTradeterm")
	sDelivery=Rs("sDelivery")
	sByer=Rs("sByer")
	sAcceptby=Rs("sAcceptby")
End if
%>
</HEAD>
<BODY BGCOLOR="#CCCCCC" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<!-- ImageReady Slices (PO1.psd) --> 
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="600">
  <TR> 
    <TD height = "100" Width="100"> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1></TD>
  </TR>
</Table>
<TABLE WIDTH="650" BORDER="0" CELLSPACING="0" CELLPADDING="0" HEIGHT="141">
  <TR> 
    <TD WIDTH="20"><IMG SRC="images/spacer.gif" WIDTH=20 HEIGHT=1></TD>
    <TD WIDTH="434">&nbsp;</TD>
    <TD WIDTH="62"><IMG SRC="images/spacer.gif" WIDTH=62 HEIGHT=18></TD>
    <TD WIDTH="128"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1">&nbsp;&nbsp;<%response.write Show_date(dpodate)%></FONT></TD>
  </TR>
  <TR> 
    <TD ROWSPAN="4" WIDTH="20"><IMG SRC="images/spacer.gif" WIDTH=20 HEIGHT=8></TD>
    <TD WIDTH="434"><IMG SRC="images/spacer.gif" WIDTH=363 HEIGHT=17></TD>
    <TD COLSPAN="2"><IMG SRC="images/spacer.gif" WIDTH=190 HEIGHT=17></TD>
  </TR>
  <TR> 
    <TD ROWSPAN="3" VALIGN="TOP" WIDTH="434"><%=sTname%> <%=sEaddress%></TD>
    <TD COLSPAN="2" VALIGN="TOP" ALIGN="CENTER" HEIGHT="40"><%=sPono%></TD>
  </TR>
  <TR> 
    <TD COLSPAN="2"><IMG SRC="images/spacer.gif" WIDTH=190 HEIGHT=19></TD>
  </TR>
  <TR> 
    <TD COLSPAN="2" HEIGHT="63" VALIGN="TOP" ALIGN="CENTER"><%=sRefno%></TD>
  </TR>
</TABLE>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="650" HEIGHT="170">
  <TR> 
    <TD WIDTH="20"> <IMG SRC="images/spacer.gif" WIDTH=20 HEIGHT=1></TD>
    <TD WIDTH="123" VALIGN="TOP"> <IMG SRC="images/spacer.gif" WIDTH=123 HEIGHT=34></TD>
    <TD WIDTH=73 VALIGN="TOP" HEIGHT="34"> <IMG SRC="images/spacer.gif" WIDTH=62 HEIGHT=34></TD>
    <TD COLSPAN=4 VALIGN="TOP" HEIGHT="34"><FONT COLOR="#000000"><%=sShipto%></FONT> 
      aa</TD>
  </TR>
  <TR> 
    <TD WIDTH="20"  ROWSPAN=3> <IMG SRC="images/spacer.gif" WIDTH=20 HEIGHT=1></TD>
    <TD ROWSPAN=3 VALIGN="TOP" WIDTH="123">&nbsp; </TD>
    <TD WIDTH=73 VALIGN="TOP" HEIGHT="33"> <IMG SRC="images/spacer.gif" WIDTH=62 HEIGHT=33></TD>
    <TD COLSPAN=4 VALIGN="TOP" HEIGHT="33"><FONT COLOR="#000000"><%=sShipby%></FONT> 
      aa</TD>
  </TR>
  <TR> 
    <TD COLSPAN=2 VALIGN="TOP" HEIGHT="34"> <IMG SRC="images/spacer.gif" WIDTH=77 HEIGHT=48></TD>
    <TD VALIGN="TOP" HEIGHT="34" WIDTH="237"><FONT COLOR="#000000"><%=sPayterm%></FONT> 
      aa </TD>
    <TD HEIGHT="2" COLSPAN="2"  VALIGN="TOP" ROWSPAN="2" WIDTH="190"> 
      <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
        <TR> 
          <TD WIDTH="91" HEIGHT="32" VALIGN="TOP"><IMG SRC="images/spacer.gif" WIDTH=91 HEIGHT=32></TD>
          <TD HEIGHT="32">aa</TD>
        </TR>
        <TR> 
          <TD HEIGHT="34" VALIGN="TOP"><IMG SRC="images/spacer.gif" WIDTH=91 HEIGHT=34><FONT COLOR="#000000"> 
            </FONT></TD>
          <TD HEIGHT="34"><FONT COLOR="#000000"><%=sDelivery%>aa </FONT></TD>
        </TR>
        <TR> 
          <TD WIDTH="91" HEIGHT="33" VALIGN="TOP"><IMG SRC="images/spacer.gif" WIDTH=91 HEIGHT=33></TD>
          <TD HEIGHT="33"><FONT COLOR="#000000"><%=sbyer%> aa</FONT></TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
   <TR> 
    <TD COLSPAN=2 VALIGN="TOP" HEIGHT="51"> <IMG SRC="images/spacer.gif" WIDTH=77 HEIGHT=51></TD>
    <TD  VALIGN="TOP" HEIGHT="51" WIDTH="237"><FONT COLOR="#000000"><%=sTradeterm%></FONT> 
      aaa </TD>
  </TR>
</table>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="600">
  <TR> 
    <TD COLSPAN=17 HEIGHT="36"> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1><IMG SRC="images/spacer.gif" WIDTH=40 HEIGHT=40></TD>
  </TR>
  <TR> 
    <TD WIDTH="20"><IMG SRC="images/spacer.gif" WIDTH=20 HEIGHT=1></TD>
    <TD><% Show_Item()%> </TD>
  </TR>
</table>
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="600">
  <TR> 
    <TD height="100" COLSPAN=17> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1></TD>
  </TR>
</table>
<P>&nbsp;</P><TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="600">
  <TR> 
    <TD WIDTH="20"> <IMG SRC="images/spacer.gif" WIDTH=20 HEIGHT=1></TD>
    <TD COLSPAN=2> <IMG SRC="images/PO1_49.gif" WIDTH=106 HEIGHT=18></TD>
    <TD COLSPAN=4 VALIGN="TOP">aa </TD>
    <TD COLSPAN=2> <IMG SRC="images/PO1_51.gif" WIDTH=50 HEIGHT=18></TD>
    <TD COLSPAN=2 VALIGN="TOP">aa </TD>
    <TD COLSPAN=6> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1></TD>
  </TR>
  <TR> 
    <TD COLSPAN=2 ROWSPAN=2> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1></TD>
    <TD COLSPAN=6 VALIGN="TOP">aa </TD>
    <TD COLSPAN=4> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1></TD>
    <TD COLSPAN=4 VALIGN="TOP">aa </TD>
    <TD WIDTH="20"> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1></TD>
  </TR>
  <TR> 
    <TD COLSPAN=6> <IMG SRC="images/PO1_59.gif" WIDTH=211 HEIGHT=45></TD>
    <TD COLSPAN=9> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1><IMG SRC="images/PO1_60.gif" WIDTH="300" HEIGHT="45"></TD>
  </TR>
  <TR> 
    <TD COLSPAN=17> <IMG SRC="images/spacer.gif" WIDTH=1 HEIGHT=1></TD>
  </TR>
</TABLE>
<!-- End ImageReady Slices --> 
</BODY> 
</HTML> 
<%
	Sub Show_Item()
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color
Color="#F8D7D1"
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM ppoitem  where spono=" & sPono &" Order By sProdcode"
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	'RS.PageSize = 5
'response.write RS2("MaxItem")	
%> 
<P>&nbsp;</P><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100%" align="left" valign="top" > 
            
      <table border="0" cellspacing="1" cellpadding="0" WIDTH="630">
        <tr> 
          <TD WIDTH="230"> <IMG SRC="images/spacer.gif" WIDTH=230 HEIGHT=32></TD>
          <TD WIDTH="110"> <IMG SRC="images/spacer.gif" WIDTH=110 HEIGHT=32></TD>
          <TD WIDTH="43"> <IMG SRC="images/spacer.gif" WIDTH=43 HEIGHT=32></TD>
          <TD WIDTH="140"> <IMG SRC="images/spacer.gif" WIDTH=140 HEIGHT=32></TD>
          <TD WIDTH="140"> <IMG SRC="images/spacer.gif" WIDTH=140 HEIGHT=32></TD>
        </tr>
        <% 
Item=1
Do While Not RS.EOF %> 
        <tr> 
          <td><FONT FACE="MS Sans Serif" SIZE="2"><%=RS("sProdcode")%>&nbsp;</FONT></td>
          <td> 
            <DIV ALIGN="CENTER"><FONT FACE="MS Sans Serif" SIZE="2"><%=RS("sprodname")%></FONT></DIV>
          </td>
          <td> 
            <DIV ALIGN="CENTER"><FONT FACE="MS Sans Serif" SIZE="2"><%=RS("nQuality")%></FONT></DIV>
          </td>
          <td  cellpadding="2">
            <DIV ALIGN="CENTER"><FONT FACE="MS Sans Serif" SIZE="2"><%=RS("nPrice")%></FONT></DIV>
          </td>
          <td  cellpadding="2"><FONT FACE="MS Sans Serif" SIZE="2"><%=RS("nAmount")%></FONT></td>
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
%> 
      </table>
          </td>
        </tr>
       </table>
<%End Sub%>