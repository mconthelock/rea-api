<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!--#include file="../List.asp" -->
<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sTaxno,sPono,sEmpno,sVendcode,chk,sVendtax,sRecno,srecby,chkpo,chkstat,nqRec,nqpo
	Dim img_sTaxno	,img_sPono,img_sEmpno,img_sVendcode,img_ddate
	Dim dDate,ddate_day,ddate_month,ddate_year,d_day,d_month,d_year,sbread,Item,sprno
	Dim prcvItemCount, ppoItemCount
	Dim compareQty, compareKey
	Dim ppoCount, prcvCount

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sTaxno=Request("staxno")
	sPono=Request.QueryString("sPono")
	sRecno=Request.QueryString("sRecno")
	sRecby=Request.QueryString("sRecby")
	If IsEmpty(Request.Form("EVENT")) Then
		If  staxno <> "" Then
			sMode = "EDIT"
				Get_Tax(sTaxno)
		Else
			sTaxno=CreateTAXNO()
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Tax(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Tax(staxno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM Ptax,preceive WHERE sTaxno='" & sTaxno & "'" &_
	" and ptax.srecno=preceive.srecno"
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
'  response.write Rs("dDate")
'  response.end
	If Not Rs.EOF Then
		sTaxno = Rs("sTaxno")
		sPono = Rs("sPono")
		sRecno = Rs("sRecno")
		srecby=RS("srecby")
		dDate = Rs("dDate")
		d_day=day(ddate)
		d_month=month(ddate)
		d_year=year(ddate)
		sEmpno = Rs("sEmpno")
		sVendcode = Rs("sVendcode")
		svendtax=RS("sVendtax")
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
	document.frmName.sTaxno.focus();
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Tax 
                  Management : </font></b></font><font color="#FFFFFF" size="2"><%=sMode%> 
                  Tax </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>" name="frmName">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Tax 
                          No.</font></b></td>
                        <td width="77%"> <%if request("desc")="desc" then%> <%=sTaxno%> 
                          <%else%> 
                          <input type="text" name="sVendTax" size="50" maxlength="90" value="<%=sVendTax%>">
                          <%=img_sTaxno%> <%end if%> </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Rec 
                          No. </font></b></td>
                        <td width="77%"> 
                          <%=srecby%>
                           </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O 
                          No. </font></b></td>
                        <td width="77%"> 
                          <%=sPono%>
                          </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Date</font></b></td>
                        <td width="77%"><font face="MS Sans Serif" size="2">Day</font> 
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
                        <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">
						Receive Tax By.</font></b></td>
                        <td width="77%"> 
                          <input type="text" name="sEmpno" size="13" maxlength="12" value="<%=Session("sEmpno")%>">
                          <%=img_sEmpno%> </td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
	 		          <input type=HIDDEN name=sRecno value="<%=sRecno%>">
	 		          <input type=HIDDEN name=spono value="<%=spono%>">
	 		          <input type=HIDDEN name=sTaxno value="<%=sTaxno%>">
	 		          <input type=HIDDEN name=status value="<%=Request("status")%>">
                    <%if  Request("desc")<> "desc" then%> 
                    <input type="submit" name="Submit" value="Save">
                    <input type="Reset" name="Reset" value="Reset">
                    <%End if%> 
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;<%Show_Item()%></td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="workflow.asp">WorkFlow Manager</a></font> <font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font></p>
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
Sub Update_Tax(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sTaxno = Request("sTaxno")
		sRecno = Request("sRecno")
		sPono = Request.Form("sPono")
		'#####	change format date from mm/dd/yyyy to dd/mm/yyyy	#####
		'ddate = Request.Form("ddate_month")&"/"&Request.Form("ddate_day")&"/"&Request.Form("ddate_year")
		ddate = Request.Form("ddate_day")&"/"&Request.Form("ddate_month")&"/"&Request.Form("ddate_year")
		'######		05/01/2010	change to old format	#####
'''		ddate = Request.Form("ddate_day")&"/"&Request.Form("ddate_month")&"/"&Request.Form("ddate_year")
'		ddate = Request.Form("ddate_day")&"/"&Request.Form("ddate_month")&"/"&Request.Form("ddate_year")+543
'		ddate = Request.Form("ddate_month")&"/"&Request.Form("ddate_day")&"/"&Request.Form("ddate_year")+543
		sEmpno = Request.Form("sEmpno")
		sVendcode = Request.Form("sVendcode")
		sVendTax = Request.Form("sVendTax")

	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(sVendTax) = "") Then
		bError = True
		img_sTaxno = sErrImage
	End If

	If (Trim(sPono) = "") Then
		bError = True
		img_sPono = sErrImage
	End If

	If (Trim(sempno) = "") Then
		bError = True
		img_sempno = sErrImage
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
		  sSQL = "SELECT * FROM  PTAX WHERE sTaxno='" & sTaxno &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "pTax", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("sTaxno") = sTaxno
		Rs("sPono") = sPono
		Rs("sRecno") = sRecno
		Rs("sEmpno") = sEmpno
	   	Rs("ddate") = ddate
		RS("sVEndTax")=sVendTax
	  Rs.Update	
	  Rs.Close

	sSQL="Select * from preceive where srecno='" & srecno & "'"
'	response.write sSQL
'	response.end
	Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText	
	if Not RS.EOF then
		if RS("cStatus")<>"I" then
			RS("cStatus")="T"
		else
			RS("cStatus")="A"
		End if
		RS.Update
	End if
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
		if Not Rs.EOF then
			sPrno=RS("sRefno")
			RS("cTax")="2"
			RS.Update
		End if
		RS.Close
		
		ChkFinish()

		IF  chkpo ="0" then
		  '########## MODIFY BY DERMPAN, 7-JAN-03, [13] ##########
		  'sSQL="Select * from ppr where sprno='" & sPrno & "'"
		  'Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			'Rs("cpo")="4"
			'RS("dFinishdate")=date
			'RS.Update
			'RS.Close
		  
		  '########## COMPARE PRECEIVEITEM AND PPOITEM QUANTITY ##########
		  'Read preceiveitem
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
		  sSql = "select sum(nquality) as total from ppoitem where spono in (select distinct spono from ppo where srefno = '" & sPrNo & "' and cStatus <> '9')"
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
		  '########## END COMPARE PRECEIVEITEM AND PPOITEM QUANTITY ##########
		 
		  '########## COMPARE KEY ##########
		  'Count ppo, Add by Dermpan on 14-JAN-03 [21]
		  sSql = "select count(*) as total from ppo where srefno = '" & sPrNo & "' and cStatus <> '9'"
		  sSql = "select sum(nQuality) as total from ppoitem where spono in (select spono from ppo where srefno = '" & sPrNo & "' and cStatus <> '9')"
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
'		  sSql = "select count(*) as total from ptax, pinvoice, preceive where ptax.spono = preceive.spono and ptax.srecno = preceive.srecno and pinvoice.spono = preceive.spono and pinvoice.srecno = preceive.srecno and preceive.sprno = '" & sPrNo & "' and preceive.spono in (select spono from ppo where srefno = '" & sPrNo & "' and cstatus <> '9')"

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
		 
		 'Update cpo
		 if compareQty and compareKey then
			sSQL="Select * from ppr where sprno='" & sPrno & "'"
			Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		  
			Rs("cpo")="4"
			RS("dFinishdate")=date
			RS.Update
			RS.Close
		 end if
		  '########## END MODIFY BY DERMPAN, 7-JAN-03, [13] ##########
		End IF
	End IF
		
	  Set Rs = Nothing
if Request("status")<>"ptaxedt" then
     Response.Redirect "../workflow/waittax.asp"
else
    Response.Redirect "../workflow/finishtax.asp"
end if
'else 
'response.write "You Don't Add"
'End if
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
%>
<%
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
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nQuality"),3)%>
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
		Do while Not Rs4.EOF
					chkpo="1"
		Rs4.Movenext
		Loop
	Else
		chkpo="0"
	End if
		RS4.Close
		Set RS4=Nothing
	
End sub
%>