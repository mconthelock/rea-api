<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition,sRecno
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sInvno,sPono,sInvby,sVendcode,dPaydate,sContno,sVendinv,chk,sRecno,item,sprno
	Dim img_sInvno	,img_sPono,img_sInvby,img_sVendcode,img_ddate,img_sContno,d_payday,d_paymonth,d_payyear
	Dim dDate,ddate_day,ddate_month,ddate_year,d_day,d_month,d_year,dpaydate_day,dpaydate_month,dpaydate_year,sbread
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sInvno=Request("sInvno")
	sPono=Request.QueryString("sPono")
	If IsEmpty(Request.Form("EVENT")) Then
		If  sInvno <> "" Then
			sMode = "EDIT"
				Get_Invoice(sInvno)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Invoice(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Invoice(sInvno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM PInvoice WHERE sInvno='" & sInvno & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sInvno = Rs("sInvno")
		sPono = Rs("sPono")
		sRecno = Rs("sRecno")
		dDate = Rs("dDate")
		d_day=day(ddate)
		d_month=month(ddate)
		d_year=year(ddate)
		sInvby = Rs("sInvby")
		dPayDate = Rs("dPayDate")
		d_Payday=day(dPaydate)
		d_Paymonth=month(dPaydate)
		d_Payyear=year(dPaydate)
		sContno=Rs("sContno")
		sVendcode = Rs("sVendcode")
		sVendinv=RS("sVendinv")
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
<title>Edit Tax</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function Form_Load(){
	document.frmName.sInvno.focus();
}
-->
</SCRIPT>

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
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
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Invoice 
                  Management : Invoice Description</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>" name="frmName">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b></b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Invoice 
                          No. : </font></b></td>
                        <td width="73%"> <%=sVEndInv%> </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b></b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Rec 
                          No. : </font></b></td>
                        <td width="73%"> <%Show_Rec(sRecno)%></td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b></b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O 
                          No. : </font></b></td>
                        <td width="73%"> <%=sPono%></td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b></b></font> 
                          <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Date 
                          : </font></b></td>
                        <td width="73%"><font face="MS Sans Serif" size="2"><%=Show_Date(ddate)%></font> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b></b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Invoice 
                          By. : </font></b></td>
                        <td width="73%"> <%Show_Empno(sInvby)%> </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Pay 
                          Date : </font></b></td>
                        <td width="73%"><font face="MS Sans Serif" size="2"><%=Show_Date(dpaydate)%></font> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Contact 
                          No.</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                          </font></b></td>
                        <td width="73%"> <%=sContno%> </td>
                      </tr>
                    </table>
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;	<%Show_Item()%> </td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="workflow.asp">WorkFlow Manager</a></font> </p>
          </td>
          <td width="2%">&nbsp;</td>
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
Sub Update_Invoice(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sInvno = Request.Form("sInvno")
		sPono = Request.Form("sPono")
		ddate = Request.Form("ddate_month")&"/"&Request.Form("ddate_day")&"/"&Request.Form("ddate_year")
		sInvby = Request.Form("sInvby")
		dPaydate = Request.Form("dPaydate_month")&"/"&Request.Form("dPaydate_day")&"/"&Request.Form("dPaydate_year")
		sContno=Request.Form("sContno")
		sVendcode = Request.Form("sVendcode")

	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(sInvno) = "") Then
		bError = True
		img_sInvno = sErrImage
	End If

	If (Trim(sPono) = "") Then
		bError = True
		img_sPono = sErrImage
	End If

	If (Trim(sInvby) = "") Then
		bError = True
		img_sInvby = sErrImage
	End If

	If (Trim(sVendcode) = "") Then
		bError = True
		img_sVendcode = sErrImage
	End If

'	If (Trim(ddate_year) = "") Then
'		bError = True
'		img_ddate = sErrImage
'	End If
chkUser()
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
if chk="1" then
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  PINVOICE WHERE sInvno='" & sInvno &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "pInvoice", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("sInvno") = sInvno
		Rs("sPono") = sPono
	   	Rs("ddate") = ddate
		Rs("sInvby") = sInvby
   		Rs("dPaydate") = dPaydate
		Rs("sContno")=sContno
		Rs("sVendcode") = sVendcode
		RS("sVendinv")=sVendinv
	  Rs.Update	

	  Rs.Close
	  sSQL="Select * from ppo where spono='" & sPono & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Rs("cStatus")="2"
		RS.Update
		RS.Close
	  Set Rs = Nothing
'     Response.Redirect "pvenderall.asp"
     Response.Redirect "../workflow/workflow.asp"
else 
response.write "You Don't Add"
End if
	else
		Show_Form()
	End If	
End Sub	
Sub chkUser()
		  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
			sSQL="Select * from ppo,pprapp where spono='" & spono & "'"
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

	sSQL = "SELECT  * FROM preceiveitem  where sRecno='" & sRecno & "'" &_
	" Order By nItem"

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
						<b><font color="#FFFFFF" size="2">product Name</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Quantity</font></b></font>
					</td>
                <td  width="3%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">U/M</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">UPrice</font></b></font>
					</td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Amount</font></b></font>
					</td>
                <td  width="3%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Cur</font></b></font>
					</td>
                <td  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Comment</font></b></font>
					</td>
<!--                <td width="8%" bgcolor="#8B2718">&nbsp;<img src="../images/dot.gif">
					</td>
                <td width="9%" bgcolor="#8B2718" align="right">
					<img src="../images/dot.gif">
					</td> -->
              </tr>
<% 
counter=0
Do While Not RS.EOF
'if RS("cStatus") <> 0  Then
'End if
Counter = Counter + 1
%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nItem")%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdname")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nQuality"),3)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 3)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber((cDbl(RS("nPrice"))*cDbl(RS("nQuality"))), 3)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sCurcode")%><img src="../images/dot.gif">
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sComment")%><img src="../images/dot.gif">
						</td>
<!--                <td width="1%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
					<img src="../images/dot.gif">
						</td>
                <td width="16%" bgcolor="<%=Color%>" align="center" height="2">
<%
'			If RS("cStatus")="2" then
%> <img src="../images/istepapp.gif" width=20 height=19>
<%'Else %>
<img src="../images/istepappw.gif" width=20 height=19> 
<%			'End if
%>
					</td> -->
              </tr>
<%

Item = cInt(RS("nItem"))+1
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
'End if
RS.MoveNext
Loop
Rs.Close
	sSQL = "SELECT sCurrency,Sum(nAmount) sumamount FROM preceiveitem,pCurrency  where sRecno='" & sRecno & "'" &_
	" and pReceiveitem.sCurcode=pCurrency.sCurcode" &_
	" Group By sCurrency" &_
	" Order By sCurrency"
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<tr>
<td colspan="9" align="center">
<%	IF Not Rs.EOF then %>
	<table width="50%">
		<tr><td colspan="2" align="center" bgcolor="teal"><Font color="#FFFFFF"><B>Total</B></Font></td></tr>
		<tr><td  width="50%" align="right" bgcolor="#F1B3A7"><Font color="#1111FF"><b><%=Formatnumber(RS("sumamount"), 3)%></b></font></td><td bgcolor="#F1B3A7"><%=RS("sCurrency")%></td></tr>
	</table>
<%	End if%>
</td>
</tr>
            </table>
<DIV Align="Center">
							<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno">
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
%>
<%
Sub Show_Rec(sRecno)
	Dim Rs2,sSQL
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL="Select sRecby from preceive where srecno='" & sRecno & "'"
 	Rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	IF Not Rs2.Eof then
		Response.write RS2("sRecby")
	End if
	RS2.Close
	Set RS2=Nothing
End Sub
%>