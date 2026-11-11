<% Option Explicit %>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sInvno,sPono,sInvby,sVendcode,dPaydate,sContno,sVendinv,chk,sRecno,chkpo,chkstat,nqRec,nqpo
	Dim img_sInvno	,img_sPono,img_sInvby,img_sVendcode,img_ddate,img_sContno,d_payday,d_paymonth,d_payyear
	Dim dDate,ddate_day,ddate_month,ddate_year,d_day,d_month,d_year,dpaydate_day,dpaydate_month,dpaydate_year,sbread
	Dim Item,sprno
	Dim prcvItemCount, ppoItemCount
	Dim compareQty, compareKey
	Dim ppoCount, prcvCount
		  
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sInvno=Request("sInvno")
	sPono=Request.QueryString("sPono")
	sRecno=Request.QueryString("sRecno")
	svendcode=Request("svendcode")

	If IsEmpty(Request.Form("EVENT")) Then
		If  sInvno <> "" Then
			sMode = "EDIT"
				Get_Invoice(sInvno)
		Else
			sInvNo=CreateInvNo()
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
                  Management : <%=sMode%> Invoice </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>" name="frmName">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Invoice 
                          No. : </font></b></td>
                        <td width="73%"> 
                          <input type="text" name="sVendInv" size="50" maxlength="90" value="<%=sVendInv%>" <%=sbread%>>
                          <%=img_sInvno%> </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right" height="4">&nbsp;</td>
                        <td width="73%" height="4"><font color="#CC3333">Invoice 
                          Number Data (Character) :Not more than 12 Characters 
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b></b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receive 
                          No. : </font></b></td>
                        <td width="73%"> 
                          <%Show_Rec(sRecno)%>
                          </td>
                      </tr>
<!--                      <tr> 
                        <td width="27%" align="right">&nbsp;</td>
                        <td width="73%"><font color="#CC3333">Purchase Order Number 
                          Data (Character) :Not more than 12 Characters </font></td>
                      </tr> -->
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b></b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O 
                          No. : </font></b></td>
                        <td width="73%"> 
                          <%=sPono%>
                          </td>
                      </tr>
<!--                      <tr> 
                        <td width="27%" align="right">&nbsp;</td>
                        <td width="73%"><font color="#CC3333">Purchase Order Number 
                          Data (Character) :Not more than 12 Characters </font></td>
                      </tr> -->
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b>*</b></font> 
                          <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Date 
                          : </font></b></td>
                        <td width="73%"><font face="MS Sans Serif" size="2">Day</font> 
                          <select name="ddate_day">
                            <%    if d_day="" then d_day=day(date)
														Create_daylist(d_day) %> 
                          </select>
                          <font size="2" face="MS Sans Serif">Month</font> 
                          <select name="ddate_month">
                            <% if d_month="" then d_month=month(date)
																Create_MonthList(d_month)%> 
                          </select>
                          Year <% if d_year="" then d_year=year(date) %> 
                          <input type="text" name="ddate_year" size="8" maxlength="4" value="<%=d_year%>">
                          <%=img_ddate%> </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right">&nbsp;</td>
                        <td width="73%"><font color="#CC3333">Recieve Date . Select 
                          From Day,Month List.</font></td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Invoice 
                          By. : </font></b></td>
                        <td width="73%"> 
                          <input type="text" name="sInvby" size="12" maxlength="12" value="<%=Session("sEmpno")%>">
                          <%=img_sInvby%> </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right">&nbsp;</td>
                        <td width="73%"><font color="#CC3333">Recieve Invoice 
                          by (Employee No.)</font></td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Pay 
                          Date : </font></b></td>
                        <td width="73%"><font face="MS Sans Serif" size="2">Day</font> 
                          <select name="dPaydate_day">
                            <%    if d_Payday="" then d_Payday=day(date)
														Create_daylist(d_Payday) %> 
                          </select>
                          <font size="2" face="MS Sans Serif">Month</font> 
                          <select name="dPaydate_month">
                            <% if d_Paymonth="" then d_Paymonth=month(date)
																Create_MonthList(d_Paymonth)%> 
                          </select>
                          Year <% if d_payyear="" then d_Payyear=year(date) %> 
                          <input type="text" name="dPaydate_year" size="8" maxlength="4" value="<%=d_Payyear%>">
                          <%=img_ddate%> </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right">&nbsp;</td>
                        <td width="73%"><font color="#CC3333">Pay Date . Select 
                          From Day,Month List.</font></td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Contact 
                          No.</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                          </font></b></td>
                        <td width="73%"> 
                          <input type="text" name="sContno" size="30" maxlength="30" value="<%=sContno%>">
                          <%=img_sContno%> </td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right">&nbsp;</td>
                        <td width="73%"><font color="#CC3333">Reference Number 
                          Data (Character) :Not more than 30 Characters </font></td>
                      </tr>
                      <tr> 
                        <td width="27%" align="right">&nbsp;</td>
                        <td width="73%"><font color="#CC0000"><b>*</b></font> 
                          Required Field.</td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name=sInvno value="<%=sInvno%>">
                    <input type=HIDDEN name=sRecno value="<%=sRecno%>">
                    <input type=HIDDEN name=spono value="<%=spono%>">
                    <input type=HIDDEN name=status value="<%=Request("status")%>">
                    <%if  Request("desc")<> "desc" then%> 
                    <input type="submit" name="Submit" value="Save">
                    <input type="submit" name="Submit2" value="Reset">
                    <%end if%> <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; <%show_Item()%></td>
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
<br></body>
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
		sInvno = Request("sInvno")
		sPono = Request.Form("sPono")
		sRecno = Request.Form("sRecno")
		'#####	04/08/2009 change format mm/dd/yyyy to dd/mm/yyyy	#####????
		ddate = Request.Form("ddate_month")&"/"&Request.Form("ddate_day")&"/"&Request.Form("ddate_year")
		'######		05/01/2010	change to old format	#####
		'''ddate = Request.Form("ddate_day")&"/"&Request.Form("ddate_month")&"/"&Request.Form("ddate_year")
		'ddate =Request.Form("ddate_day")&"/"&Request.Form("ddate_month")&"/"&Request.Form("ddate_year")+543
		sInvby = Request.Form("sInvby")
		dPaydate = Request.Form("dPaydate_day")&"/"&Request.Form("dPaydate_month")&"/"&Request.Form("dPaydate_year")
		'dPaydate = Request.Form("dPaydate_day")&"/"&Request.Form("dPaydate_month")&"/"&Request.Form("dPaydate_year")+543
		sContno=Request.Form("sContno")
		sVendcode = Request.Form("sVendcode")
		sVendInv = Request.Form("sVendInv")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(sVendInv) = "") Then
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


'	If (Trim(ddate_year) = "") Then
'		bError = True
'		img_ddate = sErrImage
'	End If
'chkUser()
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
'if chk="1" then
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  PINVOICE WHERE sInvno='" & sInvno &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "pInvoice", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
'		Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("sInvno") = sInvno
		Rs("sVendinv") = sVendInv
		Rs("sPono") = sPono
		Rs("sRecno") = sRecno
	   	Rs("ddate") = ddate
		Rs("sInvby") = sInvby
 		Rs("dPaydate") = dPaydate
		Rs("sContno")=sContno
	  Rs.Update	
	  Rs.Close

	  sSQL="Select * from pReceive where sRecno='" & sRecno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	if Not Rs.EOF then
		if RS("cStatus")<>"T"  then
			Rs("cStatus")="I"
		else
			Rs("cStatus")="A"
		End if
	End if
		RS.Update
		RS.Close
		
	sSQL="Select  sum(nQuality) sumpoitem from ppoitem where spono='" & sPono & "'"
	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
	IF Not RS.EOF Then
		nqpo=RS("sumpoitem")
	End IF
	
	RS.Close
	
	sSQL="Select  sum(nQuality) sumRecitem from pReceiveitem,preceive where pReceive.sRecno=preceiveitem.sRecno and preceive.spono='" & sPono & "'"
	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
	IF Not RS.EOF Then
		nqRec=RS("sumRecitem")
	End IF
	
	RS.Close
		
	if cInt(nqpo)=cInt(nqRec) then
		chkstat="1"
	Else	
		chkstat="0"	
	End if

	IF chkstat="1" then
	  sSQL="Select * from ppo where spono='" & sPono & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  
		sPrno=RS("sRefno")
		Rs("cInvoice")="2"
		RS.Update
		RS.Close

		ChkFinish()

		IF chkpo ="0" then
		  '########## MODIFY BY DERMPAN, 26-DEC-02, [13] ##########
		  'sSQL="Select * from ppr where sprno='" & sPrno & "'"
		  'Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		  
			'Rs("cpo")="4"
			'RS("dFinishdate")=date
			'RS.Update
			'RS.Close
		  
		  '########## COMPARE PRECEIVEITEM AND PPOITEM QUANTITY ##########
		  'Read preceiveitem, Modify by Dermpan on 14-JAN-03 [20]
		  sSql = "select sum(nquality) as total " _
		  & "from preceiveitem " _
		  & "where " _
		  & "srecno in (select distinct srecno from preceive where sprno = '" & sPrNo & "' " _
		  & "and sPoNo in (select distinct sPoNo from ppo where cStatus <> '9'))"
		  'response.write sSql & "<br>"
		  rs.open sSql, conn
		  
		  if not rs.EOF then
		  	if isNull(rs("total")) then
			  	prcvItemCount = 0
			else
			  	prcvItemCount = cInt(rs("total"))
			end if
		else
			 prcvItemCount = 0
		 end if
		 
		 rs.close

		'Read ppoitem
		  sSql = "select " _
		  & "sum(nquality) as total " _
		  & "from " _
		  & "ppoitem " _
		  & "where " _
		  & "spono in (select distinct spono from ppo where srefno = '" & sPrNo & "' and cStatus <> '9')"
		  rs.open sSql, conn
		  
		  if not rs.EOF then
		  	if isNull(rs("total")) then
				ppoItemCount = 0
			else
			  	ppoItemCount = cInt(rs("total"))
			end if
		else
			 ppoItemCount = 0
		 end if
		 
		 rs.close

		 if prcvItemCount = ppoItemCount and prcvItemCount <> 0 then
		 	compareQty = true
		 else
		 	compareQty = false
		 end if
		  '########## COMPARE PRECEIVEITEM AND PPOITEM QUANTITY ##########
		 
		  '########## COMPARE KEY ##########
		  'Count ppo, Add by Dermpan on 14-JAN-03 [21]
		  sSql = "select sum(nQuality) as total from ppoitem where spono in (select spono from ppo where srefno = '" & sPrNo & "' and cStatus <> '9')"
		  'sSql = "select count(*) as total from ppo where srefno = '" & sPrNo & "' and cStatus <> '9'"
		  'response.write sSql & "<br>"
		  rs.open sSql, conn
		  
		  if not rs.EOF then
		  	if isNull(rs("total")) then
			  	ppoCount = 0
			else
			  	ppoCount = cInt(rs("total"))
			end if
		else
			 ppoCount = 0
		 end if
		 
		 rs.close

		'Count preceive
		  sSql = "select sum(nQuality) as total " _
		  & "from " _
		  & "preceiveitem " _
		  & "where " _
		  & "sRecNo in (" _
		  & "select pReceive.sRecNo " _
		  & "from " _
		  & "ptax, pinvoice, preceive " _
		  & "where " _
		  & "ptax.srecno = preceive.srecno " _
		  & "and pinvoice.srecno = preceive.srecno and preceive.sprno = '" & sPrNo & "' " _
		  & "and preceive.spono in (select spono from ppo where srefno = '" & sPrNo & "' and cStatus <> '9')" _
		  & ")"
		  'response.write sSql & "<br>"
		  rs.open sSql, conn
		  
		  if not rs.EOF then
		  	if isNull(rs("total")) then
				prcvCount = 0
			else
			  	prcvCount = cInt(rs("total"))
			end if
		else
			 prcvCount = 0
		 end if
		 
		 rs.close

		 if prcvCount = ppoCount and prcvCount <> 0 then
		 	compareKey = true
		 else
		 	compareKey = false
		 end if
		  '########## END COMPARE KEY ##########
		 
		 '########## COMPARE P/O QUANTITY COUNT WITH P/R ITEM QUANTITY, Add on 17-FEB-03 by Dermpan ##########
		 '########## END COMPARE P/O QUANTITY COUNT WITH P/R ITEM QUANTITY ##########
		 
		 'Update cpo
		 if compareQty and compareKey then
			sSQL="Select * from ppr where sprno='" & sPrno & "'"
			Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		  
			Rs("cpo")="4"
			RS("dFinishdate")=date
			RS.Update
			RS.Close
		 end if
		  '########## END MODIFY BY DERMPAN, 26-DEC-02 ##########
		End IF
	End IF
	
	  Set Rs = Nothing
'	  response.end
'     Response.Redirect "pvenderall.asp"
if Request("status") <> "pinvedt" then
     Response.Redirect "../workflow/waitInvoice.asp"
else
   Response.Redirect "../workflow/finishInv.asp"
End if
'else 
'response.write "You Don't Add"
'End if
	else
		Show_Form()
	End If	
End Sub	
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
<%	Sub Show_Item()

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
						<b><font color="#FFFFFF" size="2">Product Name</font></b></font>
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
if Not RS.EOF then
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
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nQuality"),4)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nPrice"), 4)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber((cDbl(RS("nPrice"))*cDbl(RS("nQuality"))), 4)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sCurcode")%><img src="../images/dot.gif">
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sComment")%><img src="../images/dot.gif">
						</td>
<!--                <td width="1%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
					<img src="../images/dot.gif">
						</td>
                <td width="16%" bgcolor="<%=Color%>" align="center" height="2"><img src="../images/dot.gif"> -->
<%
'			If RS("cStatus")="2" then
%><!-- <img src="../images/istepapp.gif" width=20 height=19>
<%'Else %>
<img src="../images/istepappw.gif" width=20 height=19> -->
<%			'End if
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
'End if
RS.MoveNext
Loop
	
else
	response.write "<center><a href ='recdel.asp?sRecno="&sRecno&"'> Delete Receive </a></center>"
End if
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
		<tr><td  width="50%" align="right" bgcolor="#F1B3A7"><Font color="#1111FF"><b><%=Formatnumber(RS("sumamount"), 4)%></b></font></td><td bgcolor="#F1B3A7"><%=RS("sCurrency")%></td></tr>
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

Sub ChkFinish()
Dim chkrec,RS4,sSQL2
		Set Rs4 = Server.CreateObject("ADODB.RecordSet")
				sSQL2="Select * from ppo where  sRefno='" & sPrno & "' and (cInvoice='0' or cTax='0')"  
		Rs4.Open sSQL2, conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If Not Rs4.EOF then
'		Do while Not Rs4.EOF
					chkpo="1"
'		Rs4.Movenext
'		Loop
	Else
		chkpo="0"
	End if
		RS4.Close
		Set RS4=Nothing
	
End sub
%>