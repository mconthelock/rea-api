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
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim srecno,sprno,drecdate,srecby,sreceiver,d_day,d_month,d_year
	Dim img_srecno,sbread,Empno,chk
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
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
				sRecno=dYear & dMonth & dDay & dHour & dMinute & dSecond
			sMode = "ADD"
				sReceiver=Session("sEmpno")
		End If	
		Show_Form()
	Else
		Update_Receive(Request.Form("EVENT"))
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Receive 
                  Management : </font></b></font> <font color="#FFFFFF" size="2"><%=sMode%> 
                  Receive</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="21%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receive 
                          No. </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SRECNO" size="13" maxlength="12" value="<%=SRECNO%>" <%=sbread%>>
                          <%=img_srecno%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                          No. </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPRNO" size="13" maxlength="12" value="<%=SPRNO%>" <%=sbread%>>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receive 
                          Date </font></b></td>
                        <td width="79%">Day 
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
                          <input type="text" name="drecdate_year" size="8" maxlength="4" value="<%=d_year%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receive 
                          By </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SRECBY" size="13" maxlength="12" value="<%=SRECBY%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receiver</font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SRECEIVER" size="13" value="<%=sreceiver%>" maxlength="12">
                        </td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
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
  	sErrImage = "<IMG SRC='.../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date

		srecno = Request.Form("srecno")
		sprno = Request.Form("sprno")
		drecdate = Request.Form("drecdate_month")&"/"&Request.Form("drecdate_day")&"/"&Request.Form("drecdate_year")
		srecby= Request.Form("srecby")
		sreceiver= Request.Form("sreceiver")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(srecno) = "") Then
		bError = True
		img_srecno = sErrImage
	End If

chkUser()
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
if chk="1" then
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
		Rs("drecdate") = drecdate
		Rs("srecby") = srecby
		Rs("sreceiver") = sreceiver
	  Rs.Update	

	  Rs.Close
	  Set Rs = Nothing
     Response.Redirect "preceivedesc.asp?srecno=" & srecno
		Else
					response.write "You Don't Receive"
		end if
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