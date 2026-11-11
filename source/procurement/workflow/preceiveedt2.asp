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
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	Dim srecno,sprno,drecdate,srecby,sreceiver,d_day,d_month,d_year
	Dim invdate,inv_day,inv_month,inv_year
	Dim bldate,bl_day,bl_month,bl_year
	Dim img_srecno,sbread,Empno,chk,Item,counter,spono,chkRec,sRecold
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,img_spono
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
srecno=Request("srecno")

	If IsEmpty(Request.Form("EVENT")) Then
		If  srecno <> "" Then
			sMode = "EDIT"
			Get_Receive(srecno)
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
				sRecNO=CreateRECNO()
			sMode = "ADD"
				sPrno=Request("sPrno")
				sReceiver=Session("sEmpno")
		End If	
		Show_Form()
	Else
		sMode=Request("EVENT")
		Update_Receive(Request("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Receive(srecno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM preceive WHERE srecno='" & srecno & "'" 
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		srecno=RS("srecno")
		sprno = Rs("sprno")
		drecdate=RS("drecdate")
		d_day=day(drecdate)
		d_month=month(drecdate)
		d_year=year(drecdate)
		inv_day=day(RS("invdate"))
		inv_month=month(RS("invdate"))
		inv_year=year(RS("invdate"))
		bl_day=day(RS("bldate"))
		bl_month=month(RS("bldate"))
		bl_year=year(RS("bldate"))
		srecby=RS("srecby")
		srecold=RS("srecby")
		sreceiver = Rs("sreceiver")
		sbread="ReadOnly"
	Else
		sMode = "ADD"
		sbread="ReadOnly"
	End If
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
<title>Edit Receive :: Computer</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Receive 
                  Management :</font></b></font></b></font> <font color="#FFFFFF" size="2"><%=sMode%> 
                  Receive</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">

                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O 
                          No. </font></b></td>
                        <td width="77%"><%=Request("SPONO")%>
<!--   <input type="text" name="SPONO" size="15" maxlength="12" value="<%=Request("SPONO")%>">  -->
                          
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                          No. </font></b></td>
                        <td width="77%"> 
                          <%=SPRNO%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receieve 
                          Date </font></b></td>
                        <td width="77%">วันที่ 
                          <select name="drecdate_day">
                            <%    if d_day="" then d_day=day(date)
														Create_daylist(d_day) %> 
                          </select>
                          เดือน 
                          <select name="drecdate_month">
                            <% if d_month="" then d_month=month(date)
																Create_MonthList(d_month)%> 
                          </select>
                           คศ. <% if d_year="" then d_year=year(date) %> 
                          <input type="text" name="drecdate_year" size="8" maxlength="4" value="<%=d_year%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">* Receive 
                          No </font></b></td>
                        <td width="77%"> 
                          <input type="text" name="SRECBY" size="21" maxlength="90" value="<%=SRECBY%>"><%=Img_sRecno%>
                        </td>
                      </tr>
					  <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Invoice 
                          Date </font></b></td>
                        <td width="77%">วันที่ 
                          <select name="invdate_day">
                            <%    if inv_day="" then inv_day=day(date)
														Create_daylist(inv_day) %> 
                          </select>
                          เดือน 
                          <select name="invdate_month">
                            <% if inv_month="" then inv_month=month(date)
																Create_MonthList(inv_month)%> 
                          </select>
                          คศ. <% if inv_year="" then inv_year=year(date) %> 
                          <input type="text" name="invdate_year" size="8" maxlength="4" value="<%=inv_year%>">
                        </td>
                      </tr>
					  			                        <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">BL 
                          Date </font></b></td>
                        <td width="77%">วันที่ 
                          <select name="bldate_day">
                            <%    if bl_day="" then bl_day=0
														Create_daylist(bl_day) %> 
                          </select>
                          เดือน 
                          <select name="bldate_month">
                            <% if bl_month="" then bl_month=0
																Create_MonthList(bl_month)%> 
                          </select>
                          คศ. 
                          <input type="text" name="bldate_year" size="8" maxlength="4" value="<%=bl_year%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receiver</font></b></td>
                        <td width="77%"> 
                          <input type="text" name="SRECEIVER" size="13" value="<%=sreceiver%>" maxlength="12">
                        </td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name=sRecno value="<%=sRecno%>">
                    <input type=HIDDEN name=sRecold value="<%=sRecold%>">
                    <input type=HIDDEN name=sprno value="<%=sprno%>">
                    <input type=HIDDEN name=sPono value="<%=Request("sPono")%>"> 
                    <input type="submit" name="Submit" value="Save">
                    <input type="submit" name="Submit2" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
            <br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
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
Sub Update_Receive(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date

		srecno = Request("srecno")
		sprno = Request("sprno")
		spono = Request("spono")
		'drecdate = Request.Form("drecdate_month")&"/"&Request.Form("drecdate_day")&"/"&Request.Form("drecdate_year")
		drecdate = Request.Form("drecdate_day")&"/"& Request.Form("drecdate_month") &"/"&Request.Form("drecdate_year")&" "&Hour(Now)&":"&Minute(Now)&":"&Second(Now)
		invdate = Request.Form("invdate_day")&"/"&Request.Form("invdate_month")&"/"&Request.Form("invdate_year")&" 00:00:00"
		if Request.Form("bldate_day") = 0 or Request.Form("bldate_month") = 0 or Request.Form("bldate_year") = "" then
			bldate = ""
		else
			bldate = Request.Form("bldate_day")&"/"&Request.Form("bldate_month")&"/"&Request.Form("bldate_year")&" 00:00:00"
		end if
		srecby= Request.Form("srecby")
		srecold= Request("srecold")
		sreceiver= Request.Form("sreceiver")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	if sRecold<>srecby then
		chkRecno()
	End if

	If (Trim(srecby) = "") Then
		bError = True
		img_srecno = sErrImage
	End If
	If (Trim(spono) = "") Then
		bError = True
		img_spono = sErrImage
	End If

'chkUser()
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
'if chk="1" then
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  Preceive WHERE srecno='" & srecno &"'" 
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "Preceive", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("srecno") = srecno
		Rs("sprno") = sprno
		Rs("spono") = spono
		Rs("drecdate") = drecdate
		Rs("invdate") = invdate
		if bldate <> "" then
			Rs("bldate") = bldate
		end if
		Rs("srecby") = srecby
		Rs("sreceiver") = sreceiver
		Rs("skeyby") = Session("sEmpno")
		RS("cstatus")="0"
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing

     Response.Redirect "preceivedesc3.asp?srecno=" & srecno &"&spono=" & Request("spono")
'	Else
'				response.write "You Don't Receive"
'		end if
	else
		Show_Form()
	End If	' if Error 
End Sub	
Sub chkUser()
		  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
			sSQL="Select * from pprapp where sprno='" & sprno & "'"
		  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			If Not RS2.EOF then
				Do While Not RS2.EOF
					If RS2("sEmpno")=Session("sEmpno") then
						chk="1"
					End If
				RS2.MoveNext
				Loop
				RS2.Close
			Else
				chk="0"
			End if
End Sub
%>
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

	sSQL = "SELECT  * FROM ppritem  where sPrno="&sPrno&" Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

	'RS.PageSize = 5
'response.write RS2("MaxItem")	
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
%>

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
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Price</font></b></font>
					</td>
                <td  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Use For</font></b></font>
<!--					</td>
                <td width="8%" bgcolor="#8B2718">&nbsp;
					</td> -->
                <td width="30%" bgcolor="#8B2718" align="right" colspan="2">
						
            <DIV Valign="top"> </DIV>
					</td>
              </tr>
<% 
counter=0
Do While Not RS.EOF
if RS("cStatus") <> 0 and RS("cStatus") <> "2" Then
'End if
Counter = Counter + 1
%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nItem")%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdname")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nQuality")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nPrice")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUseFor")%>
						</td>
      <!--          <td width="1%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
						</td> -->
                <td width="16%" bgcolor="<%=Color%>" align="center" height="2" colspan="2"> 
						<input type="radio" name='<%=counter%>' value="2"> Receive 
						<input type="radio" name='<%=counter%>' value="1" Checked> No
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
							<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno">
							<INPUT type="hidden" Value='<%=counter%>' name="counter">
							<INPUT type="hidden" Value='<%=Counter%>' name="Count_Item">

                     <input type="submit" name="Submit2" value="Receive">

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
%>
<%
Sub chkRecno()
    Dim  Rs3,sSQL3,amount
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL3="Select sRecby from pReceive where sRecby='" & sRecby & "'" 
	Rs3.Open sSQL3, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if Not Rs3.EOF then
		chkRec=1
	else
		chkRec=0
	End if
	Rs3.Close
	Set RS3=nothing
End Sub

%>