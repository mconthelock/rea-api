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
'on Error resume Next
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim spono,sPrno,sexptype,dpodate,srequester,skey,nstep,Quality,Price,Amount
	Dim img_spono,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Pono
		Item=1
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
sPrno=Request("sPrno")
	IF Request("sPrno") <> "" Then 
'				Session("sExptype")=Mid(Request("sExptype"),1,2)
'	End IF
	
'	IF Request("sRequester") <> "" Then
'		Session("sRequester") = Request("sRequester")
'	End IF
			if IsEmpty(Request("sPono"))  Then
'				Session("sKey")=Session("sEmpno")
				'Session("sExptype")=Request("sexptype")
				sMode="ADD"	
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
				Pono=dYear & dMonth & dDay & dHour & dMinute & dSecond
				sPono=Pono
				Show_Form()
else
'		if Request("sPrno") <> "" and Request("Event") ="EDIT"  then
'			Session("sPrno")=Request("sPrno")
'		end if
 'If Not IsEmpty(Session("sRequester")) and Request("sRequester")<> ""  then
					Update_Ppr("ADD")
					Show_Form()
End If		
Else
		sMode="EDIT"
		Get_Ppr(Session("sPrno"))
		Show_Form()

End if
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".

Sub Get_Ppr(sprno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM ppr WHERE sprno='" & sprno & "'" &_
	"and sexptype='" & sexptype & "'"
   Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		Session("sprno")=RS("sprno")
		Session("sexptype")=RS("sexptype")
		dprdate=RS("dprdate")
		d_day=Day(dprdate)
		d_month=Month(dprdate)
		d_year=Year(dprdate)+543
		Session("srequester")=RS("srequester")
		Sesion("skey")=RS("skey")
		cprtype=RS("cprtype")
		nstep=RS("nstep")
		capprove=RS("capprove")
		sbread="ReadOnly"
	Else
		sMode = "ADD"
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
<title>PURCHASE ORDER</title>
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
          <td width="2%" height="113">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="113"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td width="79%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> 
                  PURCHASE ORDER</font></b></font></td>
                <td width="21%" bgcolor="#FF9999" align="center"> 
                  <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">DATE</font></b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"></font></div>
                </td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="93%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="4" rowspan="2"> 
                          <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Attention</b> 
                            : </font></p>
                          <p>&nbsp;</p>
                          <p>&nbsp;</p>
                          <p>&nbsp;</p>
                          <p>&nbsp;</p>
                        </td>
                        <td width="17%" height="95"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Purchase 
                          Order No. </b></font></td>
                        <td width="17%" height="95"><%=sPono%></td>
                      </tr>
                      <tr> 
                        <td width="17%" height="83"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Reference 
                          No. </b></font></td>
                        <td width="17%" height="83"><%=sPrno%> </td>
                      </tr>
                      <tr> 
                        <td colspan="3" rowspan="6"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Shipping 
                          mark </b> &nbsp;</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font></td>
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Ship 
                          To: </b></font><font color="#FFCCCC">
                          <input type="text" name="sShipto" value="" size="45" maxlength="90">
									</font> </td>
                      </tr>
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Ship 
                          By:</b>
                          <input type="text" name="sShipby" value="" size="45" maxlength="90">
								</font></td>
                      </tr>
                      <tr> 
                        <td width="51%" rowspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Terms 
                          Of Payment:</b>
                          <input type="text" name="sPayterm" value="" size="45" maxlength="45">
									</font></td>
                        <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Currency: 
                          </b></font></td>
                      </tr>
                      <tr> 
                        <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Delivery:</b>
                          <input type="text" name="sDelivery" value="" size="30" maxlength="30">
								</font></td>
                      </tr>
                      <tr> 
                        <td width="51%" rowspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Terms 
                          Of Trade:</b>
                          <input type="text" name="sTradeterm" value="" size="45" maxlength="45">
									</font></td>
                        <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Buyer 
                          Code:</b>
                          <input type="text" name="sByer" value="" size="12" maxlength="12">
									</font></td>
                      </tr>
                      <tr> 
                        <td colspan="2">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="3">&nbsp;</td>
                        <td width="51%">&nbsp;</td>
                        <td width="17%">&nbsp;</td>
                        <td width="17%">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="3">&nbsp;</td>
                        <td width="51%"><font color="#FFCCCC">...</font></td>
                        <td width="17%"><font color="#FFCCCC">...</font> </td>
                        <td width="17%"><font color="#FFCCCC">...</font> </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;<% Show_Item()%></td>
              </tr>
            </table>
            <p><br>
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()"><font size="1">Back</font></a> 
              <font size="1">|</font> <a href="../datamanager/datamanager.asp"><font size="1">Data 
              Manager</font></a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
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

<% End Sub 

	Sub Show_Item()
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color
Color="#F8D7D1"
	'sexptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	'sScriptPath = Request.ServerVariables("SCRIPT_NAME")

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

          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">

              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Item</font></b></font></td>
                <td  width="15%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Code</font></b></font></td>
                <td  width="40%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Product Name</font></b></font></td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Price</font></b></font></td>
                <td  width="20%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Amount</font></b></font></td>
  <!--              <td width="5%" bgcolor="#8B2718">&nbsp;</td>
                <td width="5%" bgcolor="#8B2718" align="right" valign="top"><img src="../Pexpenstype/images/y05.jpg" width="12" height="12"></td>-->
              </tr>
<% Do While Not RS.EOF %>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nitem")%>&nbsp;</td>
                <td width="9%" bgcolor="<%=Color%>" align="center" height="2"> <%=RS("sProdcode")%></td> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sprodname")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nQuality")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nPrice")%></td>
                <td  bgcolor="<%=Color%>" height="2">
								<%	
										Quality=RS("nQuality")
										Price=Rs("nPrice")
										Amount=Cdbl(Quality) * Cdbl(Price)
										Response.write Amount	
							%>
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

</tr>
            </table>

          </td>

        </tr>
 
      </table>

<%
End Sub
Sub Update_Ppr(sUpdateMode)
	Dim sErrImage, Rs, sSQL,dprdate
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
dprdate=Date
cprtype=Request("cprtype")
nstep=Request("nstep")
capprove=Request("capprove")

	' Check if the End date is a valid date provided it is not empty

'	If (Trim(sprno) = "") Then
'		bError = True
'		img_sprno = sErrImage
'	End If
'	If (Trim(nstep) = "") Then
'		bError = True
'		img_step = sErrImage
'	End If
'	If (Trim(Request("dprdate_year")) = "") Then
'		bError = True
'		img_dprdate_year = sErrImage
'	End If
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  ppr WHERE sprno='" & sprno &"'" &_
			"and sexptype='" & sexptype & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "ppr", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
if Prno="" then Prno=sPrno
		Rs("sprno") = Session("sPrno")
		Rs("sexptype") = Session("sexptype")
		Rs("dprdate") = dprdate
		Rs("srequester") = Session("srequester")
		Rs("skey") =Session("skey")
		Rs("cprtype") =cprtype
		Rs("nstep") =nstep
		Rs("capprove") =capprove
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
     Response.Redirect "ppr.asp?sexptype=" & sexptype & "&sKey=" & sKey & "&sprno=" & sPrno 
	else
		Show_Form()
	End If	' if Error 
End Sub	
%>