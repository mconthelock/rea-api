<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<!--#include file="scm.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sprecno,sprno,drecdate,skeyby,sreceiver,d_day,d_month,d_year
	Dim img_srecno,sbread,Empno,chk,Item,counter,spono
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim precdate,prectime, pplace
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
sprecno=Request("sprecno")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	If IsEmpty(Request.Form("EVENT")) Then
		If  sprecno <> "" Then
			sMode = "EDIT"
			Get_Receive(sprecno)
		Else
			dDate=Now
				dDay=(Day(dDate))
				If dDay < 10 then dDay="0" & Day(dDate)
						dMonth=Month(dDate)
				If dMonth < 10 then dMonth="0" & Month(dDate)
						dYear=Mid(Year(dDate),2,2)
						dHour=Hour(dDate)
				If dHour < 10 then dHour="0" & Hour(dDate)
						dMinute=Minute(dDate)
				If dMinute < 10 then dMinute="0" & Minute(dDate)
						dSecond=Second(dDate)
				If dSecond < 10 then dSecond="0" & Second(dDate)
'				sRecno=dYear & dMonth & dDay & dHour & dMinute & dSecond
				sprecno=CreatePLANNo()
				'response.write "ssssssssssss"&sprecno
				sMode = "ADD"
				sPrno=Request("sPrno")
				skeyby=Session("sEmpno")
				
		End If	
		Show_Form()
	Else
		Update_Receive(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Receive(sprecno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM PLANRECEIVE WHERE ID='" & sprecno & "'" 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sprecno=RS("ID")
		sprno = Rs("sprno")
		spono = Rs("spono")
		precdate=RS("PRECDATE")
		d_day=day(precdate)
		d_month=month(precdate)
		d_year=year(precdate)
		prectime=RS("prectime")
		pplace = Rs("precplace")
		skeyby = Rs("skeyby")
		sbread="ReadOnly"
	Else
		sMode = "ADD"
		sbread="ReadOnly"
	End If
	RS.Close
End Sub	
%>	

<%
' This routines presents the form to the user. Depending on the mode, the
' form fields are filled with data or blank.
Sub Show_Form()
	' Set the title of the page depending on if an error occurred during
	' form submission or not
%>
<html>
<head>
<title>Edit Plan Receive :: Computer</title>
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
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top"><b><img src="../images/plus.gif" width="9" height="9"> 
            WORK FLOW MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" height="265">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="265"> 

            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"></font></b></font></b></font> <font color="#FFFFFF" size="2"><%=sMode%> 
                  Plan Receive</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> <!--  <form method="post" action="<%=sScriptPath%>"> --> 
                  <form method="post" action="../workflow/Receive.asp">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                      <td width="25%" align="right"><b></b></td>
                        <td width="75%"> <!--                <input type="text" name="SRECNO" size="14" maxlength="12" value="<%'=SRECNO%>" <%'=sbread%>> --> 
                          </td>
                      </tr> 
					    <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receive 
                          No. </font></b></td>
                        <td width="75%"> <%=sprecno%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                          No. </font></b></td>
                        <td width="75%"> 
                          <%=SPRNO%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O 
                          No. </font></b></td>
                        <td width="75%"> 
                          <%=SPONO%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receieve 
                          Date </font></b></td>
                        <td width="75%"><%=Show_date(precdate)%>
<!--Day 
                          <select name="drecdate_day">
                            <%    if d_day="" then d_day=day(date)
														Create_daylist(d_day) %> 
                          </select>
                          Month 
                          <select name="drecdate_month">
                            <% if d_month="" then d_month=month(date)
																Create_MonthList(d_month)%> 
                          </select>
                          Year <% if d_year="" then d_year=year(date) %> 
                          <input type="text" name="drecdate_year" size="8" maxlength="4" value="<%=d_year%>"> -->
                        </td>
                      </tr>
                           <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Time </font></b></td>
                        <td width="75%"> <%=prectime%>
                        </td>
                      </tr>
					    <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Place </font></b></td>
                        <td width="75%"> <%=pplace%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Planer</font></b></td>
                        <td width="75%"> <%Show_Empno(skeyby)%>
<!--                          <input type="text" name="SRECEIVER" size="13" value="<%=sreceiver%>" maxlength="12">-->
                        </td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name=spRecno value="<%=spRecno%>">
                       <br>

                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2" height="2">&nbsp; <% 
'	If sMode="Show" Then
'		Show_Item()
'	Else
		Show_Item_receive()
'	End if

%><Font color="Red"> *Quantity=Plan Receive Quantity(<b>Total Quantity</b>)</font></td>
              </tr>
            </table>
            <br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
			| <a href="finishplanrec.asp">Back To Plan Receive List</a>
            | <a href="workflow.asp">WorkFlow Manager</a></font> </td>
          <td width="2%" height="265">&nbsp;</td>
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

<% End Sub %>
<%
	Sub Show_Item_Receive()
%>
<%
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color,Counter,chkreceive,chkno
Counter = 0
Color="#F8D7D1"
	sExptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  planreceive.sprno , planreceive.spono , planreceiveitem.nitem , planreceiveitem.sprodcode , planreceiveitem.sprodname , planreceiveitem.nquality as pnquality , planreceiveitem.sunitcode , ppoitem.nquality  FROM  planreceive , planreceiveitem , ppoitem  where planreceive.ID='" & spRecno & _
	"' and planreceive.ID= planreceiveitem.ID and planreceive.spono = ppoitem.spono and planreceiveitem.sprodcode = ppoitem.sprodcode" &_
	" Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

	'RS.PageSize = 5
'response.write RS2("MaxItem")	
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
%>
       <!--           <form method="post" action="../workflow/Receive.asp"> -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            
      <table width="100%" border="0" cellspacing="1" cellpadding="2">
        <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Item</font></b></font>
					</td>
                <td  width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Product Name</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Quantity</font></b></font>
					</td>
                <td  width="3%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">U/M</font></b></font>
					</td>
                <td width="5%" bgcolor="#8B2718" align="right" Valign="top" colspan="2">
						<DIV Valign="top">
							<img src="../images/dot.gif">
						</DIV>
					</td>
              </tr>
<% 
counter=0
Do While Not RS.EOF
if SumItemREC(RS("nQuality"),RS("spono"),RS("sprodcode")) <> FormatNumber(RS("nQuality"))  Then
'End if
Counter = Counter + 1
%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nItem")%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdname")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("pnQuality"))%>(<b><%=FormatNumber(sumitem(RS("pnQuality"),sprno,RS("sprodcode")))%></b>)
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%></td>
                
                <td width="5%" bgcolor="<%=Color%>" align="center" height="2" colspan="2">
<%'if RS("cStatus")<> "2" then%>
<!--						<input type="radio" name='<%=Counter%>' value="2"> Rec 
						<input type="radio" name='<%=counter%>' value="1" Checked> No -->
				<a href="../workflow/pprecitemedt2.asp?sprno=<%=sPrno%>&nitem=<%=RS("nitem")%>&skeyby=<%=skeyby%>&sRCno=<%=spRecno%>&spono=<%=Request("spono")%>"><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a>
<%'Else%>
<!--					<img src="../images/istepapp.gif" width=20 height=19>-->


					</td>
              </tr>
<%

Item = cInt(RS("nItem"))+1
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
End if
RS.MoveNext
Loop
RS.Close
%>
</td>
</tr>
            </table>

	<!--						<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno" > -->
							<INPUT type="hidden" Value='<%=counter%>' name="counter">
							<INPUT type="hidden" Value='<%=Counter%>' name="Count_Item">

          </td>
        </tr>
       </table>
</Form>
<%
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
'Response.write("<BR>nItem = "&Item)
'Response.write("<BR>Counter = "&Counter)
%>
<%
End Sub


	Sub Show_Item()

    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color,Counter,chkreceive,chkno
Counter = 0
Color="#F8D7D1"
	sExptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  * FROM ppritem  where sPrno='" & sPrno & "'" &_
	" Order By nItem"
response.write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

	'RS.PageSize = 5
'response.write RS2("MaxItem")	
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
%>
                  <form method="post" action="../workflow/Receive.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Item</font></b></font>
					</td>
                <td  width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Name</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Quantity</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">U/M</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Price</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Amount</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Currency</font></b></font>
					</td>

                <td  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Use For</font></b></font>
					</td>
                <td width="8%" bgcolor="#8B2718">&nbsp;
					</td>
                
            <td width="9%" bgcolor="#8B2718" align="right" valign="top"> 
              <DIV Valign="top"> </DIV>
					</td>
              </tr>
<% 
counter=0
Do While Not RS.EOF
if SumItemREC(nQuality,spono,sprodcode) <> FormatNumber(RS("nQuality")) Then
'End if
Counter = Counter + 1
%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nItem")%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdname")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nQuality"))%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%Show_Unit(RS("sUnitcode"))%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 3)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nAmount"))%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%Show_Currency(RS("sCurcode"))%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUseFor")%>
						</td>
                <td width="1%" bgcolor="<%=Color%>" align="center" height="2" valign="top">

						</td>
                <td width="16%" bgcolor="<%=Color%>" align="center" height="2">
<%
			If RS("cStatus")="2" then
				 Response.write "Receive"
			Else
					Response.write "No"
			End if
%>
					</td>
              </tr>
<%

Item = cInt(RS("nItem"))+1
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
End if
RS.MoveNext
Loop

%>
<tr>
<td colspan="6">&nbsp;</td>
<td width="8%" bgcolor="#FFD553" align="center" height="8"></td>
</tr>
            </table>
<DIV Align="Center">
<!--							<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno" > -->
							<INPUT type="hidden" Value='<%=counter%>' name="counter">
							<INPUT type="hidden" Value='<%=Counter%>' name="Count_Item">

 </DIV>
          </td>
        </tr>
       </table>
</Form>
<%
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
'Response.write("<BR>nItem = "&Item)
'Response.write("<BR>Counter = "&Counter)

End Sub



Function SumItemREC(nQuality,spono,sprodcode)
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select Sum(nQuality) sumitems from vprecitem where spono='" & spono &_
	"' and sprodcode='" & sProdcode	& "'"

	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText

	If Not RS2.EOF then
		if RS2("Sumitems")<>"" then
'			sumitem=cDbl(nQuality)-cDbl(RS2("sumitems"))
			SumItemREC=cDbl(RS2("sumitems"))
		else
			SumItemREC=0
		end if
	End if
	Rs2.Close
	Set Rs2=Nothing
End Function

Function SumItem(nQuality,sprno,sprodcode)
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select Sum(nQuality) sumitems from ppritem where sprno='" & sprno &_
	"' and sprodcode='" & sProdcode	& "'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText

	If Not RS2.EOF then
		if RS2("Sumitems")<>"" then
			sumitem=cDbl(RS2("sumitems"))
		else
			Sumitem=nQuality
		end if
	End if
	Rs2.Close
	Set Rs2=Nothing
End Function
%>
