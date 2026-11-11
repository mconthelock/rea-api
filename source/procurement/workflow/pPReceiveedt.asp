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
	Dim sprecno,sprno,d_day,d_month,d_year
	Dim skeyby 
	Dim img_srecno,sbread,Empno,chk,Item,counter,spono,chkRec
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,img_spono
	Dim i , precdate , precdate2 , prectime , pplace , img_pdate , img_ptime , img_place
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
		precdate=RS("PRECDATE")
		d_day=day(precdate)
		d_month=month(precdate)
		d_year=year(precdate)
		prectime=RS("prectime")
		precplace = Rs("precplace")
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Plan Receive :</font></b></font></b></font> <font color="#FFFFFF" size="2"><%=sMode%> 
                  </font></b></font></td>
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
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receieve 
                          Date </font></b></td>
                        <td width="77%">วันที่ 
                          <select name="precdate_day">
                            <%    if d_day="" then d_day=day(DateAdd("d", 1, Date))
														Create_daylist(d_day) %> 
                          </select>
                          เดือน 
                          <select name="precdate_month">
                            <% if d_month="" then d_month=month(DateAdd("d", 1, Date))
																Create_MonthList(d_month)%> 
                          </select>
                          พศ. <% if d_year="" then d_year=year(DateAdd("d", 1, Date)) %> 
                          <input type="text" name="precdate_year" size="8" maxlength="4" value="<%=d_year%>"><%=img_pdate%>	
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">* Plan Receive 
                          TIME </font></b></td>
                        <td width="77%"> 
                               <select name="ptime1" >
                                  <option value=""></option>
                                  <%
					for i=0 to 23 
						response.write "<option value='" & Right("0"&i,2) &"'>" & Right("0"&i,2) &"</option>"  & chr(10) & chr(13)
					next
				%>
                                </select>
                                : 
                                <select name="ptime2" >
                                  <option value=""></option>
                                  <%
					for i=0 to 59 
						response.write "<option value='" &  Right("0"&i,2)  &"'>" & Right("0"&i,2) &"</option>"  & chr(10) & chr(13)
					next 					
				%>
                                </select><%=img_ptime%>			
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">* Place</font></b></td>
                        <td width="77%"> 
                            <input type="radio" name="PLACE" id="PLACE1" value="1" checked>Warehouse<br>
							<input type="radio" name="PLACE" id="PLACE2" value="2">Other&nbsp;&nbsp;<input type="text" name="PLACETXT" id = "PLACETXT" value="" ><br><%=img_place%>
                        </td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name=spRecno value="<%=spRecno%>">
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
	    sprecno = Request("sprecno")
	'	precdate = Request.Form("precdate_month")&"/"&Request.Form("precdate_day")&"/"&Request.Form("precdate_year")
		precdate = Request.Form("precdate_day")&"/"&Request.Form("precdate_month")&"/"&Request.Form("precdate_year")
	
		precdate2 = Request.Form("precdate_day")&"/"&Request.Form("precdate_month")&"/"&Request.Form("precdate_year")
		sprno = Request("sprno")
		spono = Request("spono")
		prectime = Request.Form("ptime1")&":"&Request.Form("ptime2")
		skeyby = Session("sEmpno")
		
		'######		04/08/2009	change format date from mm/dd/yyyy to dd/mm/yyyy	#####
		'drecdate = Request.Form("drecdate_month")&"/"&Request.Form("drecdate_day")&"/"&Request.Form("drecdate_year")
		'######		05/01/2010	change to old format	#####
'''		drecdate = Request.Form("drecdate_day")&"/"&Request.Form("drecdate_month")&"/"&Request.Form("drecdate_year")
		'srecby= Request.Form("srecby")
		'sreceiver= Request.Form("sreceiver")
		'skeyby = Session("sEmpno")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty
	'chkRecno()
	
	
	If (Trim(Request.Form("ptime1")) = "") Then
		bError = True
		img_ptime = sErrImage
	End If
	If (Trim(Request.Form("ptime2")) = "") Then
		bError = True
		img_ptime = sErrImage
	End If
	If (Trim(Request.Form("precdate_year")) = "") Then
		bError = True
		img_pdate = sErrImage
	End If
	If (DateDiff("d",Date,precdate2) <= 0) Then
			bError = True
			img_pdate = sErrImage
	End If
	If (Trim(Request.Form("PLACE")) = "2") Then
		If (Trim(Request.Form("PLACETXT")) = "") Then
			bError = True
			img_place = sErrImage
		End If
	End If
	If (Trim(Request.Form("PLACE")) = "1") Then
		pplace = "Warehouse"
	else
	    pplace = Request.Form("PLACETXT")
	end if
'chkUser()
If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
'if chk="1" then

       If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  planreceive WHERE id='" & sprecno &"'" 
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "planreceive", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("id") = sprecno
		Rs("sprno") = sprno
		Rs("spono") = spono
		Rs("skeyby") =skeyby
		Rs("skeydate") = month(Date)&"/"&day(Date)&"/"&year(Date)
		Rs("precdate") = precdate
		RS("prectime")= prectime
		RS("precplace") = pplace 
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing 
      
     Response.Redirect "ppreceivedesc2.asp?sprecno=" & sprecno &"&spono=" & spono
	  response.end	
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
		<%if FormatNumber(sumitemrec(RS("nQuality"),RS("spono"),RS("sprodcode")))=FormatNumber(RS("nQuality"),3) then%>
			<font color="red">
		<%End if%>
				<%=FormatNumber(sumitemrec(RS("nQuality"),RS("spono"),RS("sprodcode")))%>(<b><%=FormatNumber(RS("nQuality"),3)%></b>)
		<%if FormatNumber(sumitemrec(RS("nQuality"),RS("spono"),RS("sprodcode")))=FormatNumber(RS("nQuality")) then%>
			</font>
		<%End if%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 3)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nAmount"))%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sCurcode")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sComment")%>
						</td>
                <td width="16%" bgcolor="<%=Color%>" align="center" height="2" colspan="2">
				
		<%if FormatNumber(sumitemrec(RS("nQuality"),RS("spono"),RS("sprodcode"))) <> FormatNumber(RS("nQuality")) then%>

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

Function SumItem(nQuality,spono,sprodcode)
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select Sum(nQuality) sumitems from planreceive , planreceiveitem where planreceive.id = planreceiveitem.id and spono='" & spono &_
	"' and sprodcode='" & sProdcode	& "'"

	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText

	If Not RS2.EOF then
		if RS2("Sumitems")<>"" then
'			sumitem=cDbl(nQuality)-cDbl(RS2("sumitems"))
			sumitem=cDbl(RS2("sumitems"))
		else
			Sumitem=0
		end if
	End if
	Rs2.Close
	Set Rs2=Nothing
End Function
%>

 

