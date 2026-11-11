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
on error resume next
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,chkprod,chkvender
	Dim sprno,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Prno,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode,img_cStatus
	Dim sUnitcode,sQuotation,img_prodcode,sImagespec
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,sItemno
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")

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
				   sitemno=dYear & dMonth & dDay & dHour & dMinute & dSecond

		nitem=Request("nitem")
		sprno=Request("sprno")
		sExptype = Request("sExptype")
		sVendcode = Request("sVendcode")
		sProdcode = Request("sProdcode")

		if sProdcode = "" and sVendcode = "" Then
				Session("sProdcode") = ""
				Session("sVendcode") = ""
		end if

		if sProdcode = "" Then
			sProdcode = Session("sProdcode")
'			sExptype = Session("sExptype2")
		  	sSQL = "SELECT * FROM pProduct,pCurrency,pUnit WHERE sProdcode='"&sProdcode&"' and sExptype = '"&sExptype&"'"
			Set Rs = Server.CreateObject("ADODB.RecordSet")
			Rs.MaxRecords = 1
			Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			sProddesc = RS("sEDesc")
			sProdname = RS("sEprodname")
			sVendcode = Rs("sVendcode")
			nprice=RS("nprice")
			sCurcode=Rs("sCurcode")
			sUnitcode=Rs("sUnitcode")
			sImagespec=RS("sImagespec")
		else 
			Session("sProdcode") = sProdcode
		  	sSQL = "SELECT * FROM pProduct WHERE sProdcode='"&sProdcode&"' and sExptype = '"&Request("sExptype")&"'" 
			Set Rs = Server.CreateObject("ADODB.RecordSet")
			Rs.MaxRecords = 1
			Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			sProddesc = RS("sEDesc")
			sProdname = RS("sEprodname")
			Session("sExptype2") = sExptype
			sVendcode = Rs("sVendcode")
			nprice=RS("nprice")
			sCurcode=Rs("sCurcode")
			sUnitcode=Rs("sUnitcode")
			sImagespec=RS("sImagespec")
		end if

	If IsEmpty(Request.Form("EVENT")) Then
		If  nitem <> "" Then
			sMode = "EDIT"
			if Request("status")<>"change" then
				Get_Pritem(sPrno)
			else
				sMode = "ADD"
			end if
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		sMode=Request("EVENT")
		Update_Pritem(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Pritem(sPrno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  	sSQL = "SELECT * FROM ppritem WHERE nitem='" & nitem & "'" &_
	"and sprno='" & sprno & "'"

  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sPrno=RS("sPrno")
		nItem = Rs("nItem")
		sProdcode=RS("sProdcode")
		sProdname=RS("sProdname")
		sVendcode= Rs("Svendcode")
		sProddesc=RS("sProddesc")
		nQuality=RS("nQuality")
		nPrice=RS("nPrice")
		nAmount=RS("nAmount")
		sCurcode=RS("sCurcode")
		sUseFor=RS("sUseFor")
		sQuotation=RS("sQuatation")
		sComment=RS("sComment")
		sUnitcode=RS("sUnitcode")
		cStatus=RS("cStatus")
		sImagespec=RS("sImagespec")
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
<title><%=sMode%> P/R Item</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</Script>
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../datamanager/images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../security/images/y02.jpg" width="12" height="12"></td>
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  P/R Item</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                          No. : </font></b></td>
                        <td width="75%"> <%=sPrno%></td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Item 
                          No. : </font></b></td>
                        <td width="81%"><%=nItem%></td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Code : </font></b></td>
                        <td width="81%"><%=sProdcode%>
				<%'if sMode<>"EDIT" then%>
<!--                          | <A HREF="JavaScript:OpenNewWindow('../Datamanager/SearchProduct.asp?sPrno=<%=sPrno%>&nItem=<%=nItem%>&sExptype=<%=Request("sExptype")%>&smode=<%=sMode%>&approve=app','SearchProduct','width=500,height=500')">Search 
                          Product</A> | -->
				<%'End if%>
						</td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Name : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="sProdname" size="30" value="<%=server.htmlEncode(sProdname)%>" maxlength="255">
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                          Code : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="sVendcode" size="15" maxlength="12" value="<%=sVendcode%>"><%Show_Vender(sVendcode)%>
                          <%=img_sVendcode%> </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Description : </font></b></td>
                        <td width="81%"> 
                          <textarea name="sProddesc" cols="30" rows="3"><%=sProddesc%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quantity 
                          : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="nQuality" size="14" maxlength="12" value="<%=nQuality%>">
                          <%=img_nQuality%>
                        </td>
                      </tr>
					  <tr>
						<td width="19%" align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b> U/M : </b></font></td>
						<td width="81%">
				          <select name="sUnitcode" size=1>
                            <%List_Unit(sUnitcode)%> 
                          </select> 
						</td>
					  </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UPrice 
                          : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="nPrice" size="14" maxlength="12" value="<%=nPrice%>">
                          <%=img_nPrice%>
                        </td>
                      </tr>
                      <!--                      <tr> 
                        <td width="19%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Amount</font></td>
                        <td width="81%"> 
                          <input type="text" name="nAmount" size="19" maxlength="19" value="<%=nAmount%>"><%=img_nAmount%>
                        </td>
                      </tr> --> 
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Currency 
                          : </font></b></td>
                        <td width="81%"> 
                          <SELECT NAME="sCurcode" SIZE=1>
                            <%List_Currency(sCurcode)%> 
                          </SELECT>
                        </td>
                      </tr>
<!--                       <tr> 
                        <td width="19%" align="Right" height="33"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UseFor 
                          : </font></b></td>
                        <td width="81%" height="33"> 
                          <input type="text" name="sUsefor" size="30" value="<%=sUsefor%>" maxlength="120">
                        </td>
                      </tr> -->
                      <tr> 
                        <td width="19%" align="Right" height="33"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quotation No. 
                          : </font></b></td>
                        <td width="81%" height="33"> 
                          <input type="text" name="sQuotation" size="21" value="<%=sQuotation%>" maxlength="20">
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Comment 
                          : </font></b></td>
                        <td width="81%"> 
                          <textarea name="sComment" cols="30" rows="3"><%=sComment%></textarea>
                        </td>
                      </tr>
                       <tr> 
                        <td width="19%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Image :</font></b></td>
                        <td width="81%"> 
							<a href="../datamanager/Productimages/<%=sImagespec%>" target="_blank"> <%=sImagespec%></a> &nbsp; &nbsp;
							<%if sImagespec<>"" then%>
								<input type="checkbox" name="del" value="del">Del<br>
							<%End if%>
                        </td>
                      </tr>
<!--                      <%'if sMode<>"ADD" then%> 
                                           <tr> 
                        <td width="19%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Status</font></td>
                        <td width="81%"> 
                          <input type="text" name="cStatus" size="2" maxlength="1" value="<%=cStatus%>"><%=img_cStatus%>
                        </td>
                      </tr> --> <%'end if%> 
                    </table>
                    <br>
                    <input type=HIDDEN name="sprno" value="<%=sprno%>">
                    <input type=HIDDEN name="sprodcode" value="<%=sprodcode%>">
                    <input type=HIDDEN name="nItem" value="<%=nItem%>">
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name=sImagespec value="<%=sImagespec%>">
                    <input type=HIDDEN name=sExptype value="<%=sExptype%>">
                    <input type="submit" name="Submit" value="Save">
                    <input type="Reset" name="Submit2" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="workflow.asp">WorkFlow Manager</a></font> </p>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../security/images/y04.jpg" width="12" height="12"></td>
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
Sub Update_Pritem(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sPrno = Request("sPrno")
		nItem = Request("nItem")
		sProdcode= Request.Form("sProdcode")
		sProdname= Request.Form("sProdname")
		sVendcode= Request.Form("sVendcode")
		sProddesc= Request.Form("sProddesc")
		nQuality= Request.Form("nQuality")
		nPrice= Request.Form("nPrice")
'		nAmount= Request.Form("nAmount")
		sCurcode= Request.Form("sCurcode")
		sUseFor= Request.Form("sUseFor")
		sQuotation= Request.Form("sQuotation")
		sUnitcode= Request.Form("sUnitcode")
		sComment= Request.Form("sComment")
		simagespec=Request("sImagespec")
		cStatus="9"
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

'	chkproduct()

'	If chkprod <> 1 Then
'		bError = True
'		img_Prodcode = sErrImage
'	End If
	If (Trim(nQuality) = "") Or Not IsNumeric(nQuality) Or nQuality <= 0 Then
		bError = True
		img_nQuality = sErrImage
	End If
	If (Trim(nPrice) = "") Then
'		bError = True
'		img_nPrice = sErrImage
		nPrice=0
	End If
'	If (Trim(nAmount) = "") Then
'		bError = True
'		img_nAmount = sErrImage
'	End If
	chk_vender()
	If chkvender=0 and sVendcode<>"" Then
		bError = True
		img_sVendcode = sErrImage
	End If

	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode

		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  ppritem WHERE sPrno='" & sPrno &"'" &_
			"and nitem='" & nitem & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "ppritem", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		Rs("sPrno") = sPrno
		Rs("nItem") = nItem
		Rs("sProdcode") = sProdcode
		Rs("sProdname") = sProdname
		Rs("sVendcode") =sVendcode
		Rs("sProddesc") =sProddesc
		Rs("nQuality") =nQuality
		Rs("nPrice") =FormatNumber(nPrice, 3)
		Rs("nAmount") =FormatNumber((cDbl(nQuality)*cDbl(nPrice)), 3)
		Rs("sCurcode") =sCurcode
		Rs("sUseFor") =sUseFor
		Rs("sQuatation") =sQuotation
		Rs("sUnitcode") =sUnitcode
		Rs("sComment") =sComment        
		Rs("cStatus") =cStatus 
			if Request("del")="del" then
				sImagespec="" 
			End if
		RS("sImagespec")=sImagespec          
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
'     Response.Redirect "showpr.asp?sPrno=" & sPrno & "&Event=add"
     Response.Redirect "pprapp.asp?sPrno=" & sPrno & "&Event=add"
'Response.wrie "Complete"
	ShowForm(sprno)
	else
		Show_Form()
	End If 
End Sub	
Sub chkProduct()
	Dim Rs2,sSQL2
    Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select * from pproduct where sprodcode='" & sprodcode & "'" &_
	" and sExptype='" & sExptype & "'"

	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS2.EOF then
			chkprod=1
		End if
	Rs2.Close
	Set Rs2=nothing
End Sub
%>
<%
Sub show_file()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from ppritemdesc Where sprno='" & sprno & "'" &_
		" and nitem='" & nitem & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
%><a href="../workflow/ppritemdescimages/<%=RS2("sfile")%>" target="_blank">
<%					Response.Write(Rs2("sfile") & "</a>&nbsp;&nbsp;&nbsp;")
					Response.write(RS2("sdesc") & "&nbsp;&nbsp;&nbsp;&nbsp;")
%><a Href="delfile.asp?sprno=<%=sprno%>&nitem=<%=nitem%>&sExptype=<%=Request("sExptype")%>&sitemno=<%=RS2("sitemno")%>">Del</a><br>
<%
			Rs2.MoveNext
		Loop
	End IF
End Sub
%>
<%
Sub chk_vender()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pvender Where sVendcode='" & sVendcode & "'"

		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		chkvender=1
	else
		chkvender=0
	End IF
End Sub
%>