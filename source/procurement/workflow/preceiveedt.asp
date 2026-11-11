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
	Dim srecno,sprno,drecdate,srecby,sreceiver,d_day,d_month,d_year,dinvdate,dbldate
	Dim img_srecno,sbread,Empno,chk,Item,counter,spono,chkRec
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,img_spono
	Dim skeyby
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
srecno=Request("srecno")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
  
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
	
		Update_Receive(Request.Form("EVENT"))
		'response.write Request.Form("drecdate_month")&"/"&Request.Form("drecdate_day")&"/"&Request.Form("drecdate_year") & " " & hour(now) & ":" & Minute(now) & ":" & Second(now)
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
		srecby=RS("srecby")
		sreceiver = Rs("sreceiver")
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
<title>Edit Receive :: Computer</title>
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
                        <td width="77%"> <%=Request("SPONO")%>
<!--  <input type="text" name="SPONO" size="15" maxlength="12" value="<%=Request("SPONO")%>" <%=sbread%>>  -->
                          
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
                          <input type="text" name="SRECBY" size="50" maxlength="90" value="<%=SRECBY%>"><%=Img_sRecno%>
                        </td>
                      </tr>
					   <tr>
					    <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Invoice 
                          Date </font></b></td>
                        <td width="77%">วันที่ 
                          <select name="dinvdate_day">
                            <%    if d_day="" then d_day=day(date)
														Create_daylist(d_day) %> 
                          </select>
                          เดือน 
                          <select name="dinvdate_month">
                            <% if d_month="" then d_month=month(date)
																Create_MonthList(d_month)%> 
                          </select>
                          คศ. <% if d_year="" then d_year=year(date) %> 
                          <input type="text" name="dinvdate_year" size="8" maxlength="4" value="<%=d_year%>">
                        </td>
					  </tr>
					  <tr>
					    <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">BL 
                          Date </font></b></td>
                        <td width="77%">วันที่ 
                          <select name="dbldate_day">
                            <%   
														Create_daylist(0) %> 
                          </select>
							เดือน 
                          <select name="dbldate_month">
                            <% 
																Create_MonthList(0)%> 
                          </select>
							คศ.  
                          <input type="text" name="dbldate_year" size="8" maxlength="4" value="">
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
                    <input type=HIDDEN name=sprno value="<%=sprno%>">
                    <input type="hidden" name=sPono value="<%=Request("sPono")%>">
                    <input type="submit" name="Submit" value="Save">
                    <input type="submit" name="Submit2" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">
				<%
						Show_Item_Receive()
				%>
				
&nbsp; </td>
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
		'######		04/08/2009	change format date from mm/dd/yyyy to dd/mm/yyyy	#####
'''		drecdate = Request.Form("drecdate_month")&"/"&Request.Form("drecdate_day")&"/"&Request.Form("drecdate_year") & " " & hour(now) & ":" & Minute(now) & ":" & Second(now)
		'######		05/01/2010	change to old format	#####
		drecdate = Request.Form("drecdate_day")&"/"&Request.Form("drecdate_month")&"/"&Request.Form("drecdate_year")&" "&Hour(Now)&":"&Minute(Now)&":"&Second(Now)
		dinvdate = Request.Form("dinvdate_day")&"/"&Request.Form("dinvdate_month")&"/"&Request.Form("dinvdate_year")&" 00:00:00"
		if Request.Form("dbldate_day") = 0 or Request.Form("dbldate_month") = 0 or Request.Form("dbldate_year") = ""  then
				dbldate = ""
		else
				dbldate = Request.Form("dbldate_day")&"/"&Request.Form("dbldate_month")&"/"&Request.Form("dbldate_year")&" 00:00:00"
		end if	
		srecby= Request.Form("srecby")
		sreceiver= Request.Form("sreceiver")
		skeyby = Session("sEmpno")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty
	'response.write "ttttttttttt"
	
	chkRecno()
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
		Rs("drecdate") =drecdate
		Rs("invdate") =dinvdate
		if dbldate <> "" then
			Rs("bldate") = dbldate
		end if
		Rs("srecby") = srecby
		Rs("sreceiver") = sreceiver
		RS("cstatus")="0"
		RS("skeyby") = skeyby
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
     Response.Redirect "preceivedesc2.asp?srecno=" & srecno &"&spono=" & Request("spono")
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


	Sub Show_Item_Receive()

    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color,Counter,chkreceive,chkno
	Dim TotalRows,TotalPage,gPage,Page,PageNo,PageSize,RowCount,i,MaxPage
Counter = 0
Color="#F8D7D1"
	sExptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  * FROM ppoitem  where sPono='" & Request("sPono") & "'" &_
	" Order By sprodcode"
	'response.write sSQL
	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
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
       <!--           <form method="post" action="../workflow/Receive.asp"> -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            
      <table width="104%" border="0" cellspacing="1" cellpadding="2">
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
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">U/M</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">UPrice</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Amount</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Cur</font></b></font>
					</td>

                <td  bgcolor="#8B2718" width="1%"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Comment</font></b></font>
					</td>

                <td width="30%" bgcolor="#8B2718" align="right" Valign="top" colspan="2">
						<DIV Valign="top">
							<img src="../images/dot.gif">
						</DIV>
					</td>
              </tr>
<% 
counter=0
item=1
If Not RS.EOF then
		RS.Absolutepage=PageNo
		RowCount=1
		 Do While Not RS.EOF and RowCount<=PageSize

'End if
Counter = Counter + 1
%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=((PageNo-1)*RS.PageSize)+RowCount%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdname")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2">
		<%if FormatNumber(sumitem(RS("nQuality"),RS("spono"),RS("sprodcode")))=FormatNumber(RS("nQuality"),3) then%>
			<font color="red">
		<%End if%>
				<%=FormatNumber(sumitem(RS("nQuality"),RS("spono"),RS("sprodcode")),3)%>(<b><%=FormatNumber(RS("nQuality"),3)%></b>)
		<%if FormatNumber(sumitem(RS("nQuality"),RS("spono"),RS("sprodcode")))=FormatNumber(RS("nQuality")) then%>
			</font>
		<%End if%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 4)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nAmount"))%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sCurcode")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sComment")%>
						</td>
                <td width="16%" bgcolor="<%=Color%>" align="center" height="2" colspan="2">
				
		<%if FormatNumber(sumitem(RS("nQuality"),RS("spono"),RS("sprodcode")),3) <> FormatNumber(RS("nQuality"),3) then%>

<%Else%>
					<img src="../images/istepapp.gif" width=20 height=19>
<%		End if%>

					</td>
              </tr>
<%

Item = Item+1
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
'End if
RowCount=RowCount+1
RS.MoveNext
Loop
End if
%>
        <%if TotalPage > 1 then%> 
        <form action= "<%=sScriptPath%>" method="Post">
          <tr> 
            <td colspan="9" align="center" bgcolor="#FFD553"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td align="center" bgcolor="#FFD553" width="2%"><img src="../datamanager/images/dot.gif"></td>
                  <td align="left" bgcolor="#FFD553" width="45%"> <%
			If PageNo=1 then
				response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Prev&sPrno=<%=sPrno%>&sPono=<%=Request("sPono")%>">Previous 
                    Page</a> <%End if%> 
                  </td>
                  <td align="right" bgcolor="#FFD553" width="45%"> <% if PageNo=TotalPage Or PageNo=MaxPage then
				Response.write " "
			Else
%> <A Href="<%=sScriptPath%>?PageNo=<%=PageNo%>&page=Next&sPrno=<%=sPrno%>&sPono=<%=Request("sPono")%>">Next 
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
                    <A Href="<%=sScriptPath%>?PageNo=<%=i%>&maxPage=<%=TotalPage%>&sPrno=<%=sPrno%>&sPono=<%=Request("sPono")%>"><%=i%></a> 
                    <%
	End if
	Next %></td>
                </tr>
              </table>
              <%End if%> 
            </td>
          </tr>
        </form>
<tr>
<td colspan="6">&nbsp;</td>
<td width="8%" bgcolor="#FFD553" align="center" height="8"></td>
</tr>
            </table>
<DIV Align="Center">
	<!--						<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno" > -->
							<INPUT type="hidden" Value='<%=counter%>' name="counter">
							<INPUT type="hidden" Value='<%=Counter%>' name="Count_Item">
      </DIV>
          </td>
        </tr>
       </table>

<%
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
'Response.write("<BR>nItem = "&Item)
'Response.write("<BR>Counter = "&Counter)

End Sub

Function SumItem(nQuality,spono,sprodcode)
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select Sum(nQuality) sumitems from vprecitem where spono='" & spono &_
	"' and sprodcode='" & sProdcode	& "'"



'	sSQL2="Select nvl(Sum(nQuality),0) sumitems from vprecitem where spono='" & spono &_
'	"' and sprodcode='" & sProdcode	& "'"

'response.write sSQL2
'response.end
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText

	If Not RS2.EOF then
		if RS2("Sumitems")<>"" then
'			sumitem=cDbl(nQuality)-cDbl(RS2("sumitems"))
			SumItem=cDbl(RS2("sumitems"))
		else
			SumItem=0
		end if
	End if
	Rs2.Close
	Set Rs2=Nothing
End Function
%>