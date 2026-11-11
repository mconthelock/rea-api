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
'on error resume next
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprno,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Prno,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode,img_cStatus
	Dim sUnitcode,sQuotation,sfile,sItemno,chkproduct,chkvender,chkexict,sImagespec
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	if Request("sItemno")="" then
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
	else
		sItemno=Request("sItemno")
	End if
	if Session("nItem")="" then
		nitem=Request("nitem")
	else
		nitem=Session("nItem")
	End if
	
		Session("nItem")=nItem
'		nitem=Request("nitem")
		sprno=Request("sprno")
		sExptype = Request("sExptype")
		sVendcode = Request("sVendcode")
		sProdcode = Request("sProdcode")

	'if Request("sVendcode") <> "" then
		'sSQL = "select * from pVender where sVendcode = '" & Request("sVendcode") & "' and pvender.cstatus='1'"
		'response.write sSQL
	'	Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	'	if Rs.EoF then
	'		Session("sVendcode") = ""
	'	end if
	'	Rs.Close
	'end if 


		if sProdcode = "" and sVendcode = "" Then
				Session("sProdcode") = ""
				Session("sVendcode") = ""
		end if

		if sProdcode = "" Then
			sProdcode = Session("sProdcode")
			sExptype = Session("sExptype2")
		  	sSQL = "SELECT * FROM pProduct,pCurrency,pUnit WHERE sProdcode='"&sProdcode&"' and sExptype = '"&sExptype&"'"
			
			Rs.MaxRecords = 1
			Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			sProddesc = RS("sEDesc")
			sProdname = RS("sEprodname")
			sVendcode = Rs("sVendcode")
			nprice=RS("nprice")
			sCurcode=Rs("sCurcode")
			sUnitcode=Rs("sUnitcode")
			sImagespec=Rs("sImage")
		else 		
			Session("sProdcode") = sProdcode
			Set Rs = Server.CreateObject("ADODB.RecordSet")

			
		  	sSQL = "SELECT * FROM pProduct WHERE sProdcode = '"&sProdcode&"' and sExptype = '"&Request("sExptype")&"'" 
			Rs.MaxRecords = 1
			Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			sProddesc = RS("sEDesc")
			sProdname = RS("sEprodname")
			Session("sExptype2") = sExptype
			sVendcode = Rs("sVendcode")
			nprice=RS("nprice")
			sCurcode=Rs("sCurcode")
			sUnitcode=Rs("sUnitcode")
			sImagespec=Rs("sImage")
		end if
	If IsEmpty(Request("EVENT")) Then
		If  nitem <> "" AND request("smode")="EDIT" Then
			sMode = "EDIT"
			if Request("status")<>"change" then
				Get_Pritem(sPrno)
			else
				sMode = "EDIT"
			end if
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Pritem(Request("EVENT"))
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
<meta http-equiv="Content-Type" content="text/html; charset=">
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
                        <td width="81%"> 
                          <input type="text" name="sProdcode" size="15" maxlength="12" value="<%=sProdcode%>">
                          | <A HREF="JavaScript:OpenNewWindow('../datamanager/SearchProduct.asp?sPrno=<%=sPrno%>&nItem=<%=nItem%>&sExptype=<%=Request("sExptype")%>','SearchProduct','width=600,height=500')">Search 
                          Product</A> | </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Name : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="sProdname" size="30" value="<%=sProdname%>" maxlength="120">
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                          Code : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="sVendcode" size="15" maxlength="12" value="<%=sVendcode%>"> : <%Show_Vender(sVendcode)%>
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
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><Blink><b>*</b></Blink>Quantity 
                          : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="nQuality" size="14" maxlength="12" value="<%=nQuality%>">
                          <%=img_nQuality%>                         </td>
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
                          <input type="text" name="nPrice" size="19" maxlength="12" value="<%=nPrice%>">
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
<!--                      <tr> 
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
                        <td width="81%"><a href="../datamanager/Productimages/<%=sImagespec %>" target="_blank"> <%=sImagespec %></a> &nbsp; &nbsp;
							<%if sImagespec<>"" then%>
								<input type="checkbox" name="del" value="del">Del<br>
							<%End if%>
                        </td>
                      </tr>
<!--                       <tr> 
                        <td width="19%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Attach Files :</font></b></td>
                        <td width="81%"> <%show_file() %><br>
							<A HREF="JavaScript:OpenNewWindow('ppritemdescedt.asp?sPrno=<%=sPrno%>&nItem=<%=nItem%>&sitemno=<%=sItemno%>','SearchProduct','width=500,height=500')">Add Attach Files</a>
							<%
'							Session("sprno")=sprno
'							Session("nitem")=nitem
'							Session("sitemno")=sitemno
%>
                        </td>
                      </tr> -->

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
                    <input type=HIDDEN name="nItem" value="<%=nItem%>">
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name=approve value="<%=Request("approve")%>">
                    <input type=HIDDEN name=sExptype value="<%=Request("sExptype")%>">
                    <input type=HIDDEN name=sImagespec value="<%=sImagespec%>">
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
            <hr noshade size="1">
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#666666"><b><font size="1">MITSUBISHI 
            ELEVATOR ASIA CO.,LTD.</font><br>
            </b><font size="1">700/86 Amata Nakorn Industrial Estate,Moo 6,Bangna-Trad Rd., <br>
T.Don Hua Roh,Muang District,Chonburi 20000 <br>
Tel:(038) 213-170-80 Fax: (038) 213-184-6</font></font></td>
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
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	' Get the form data into variables
	' Check to see if Start date is a valid date
		sPrno = Request("sPrno")
		nItem = Request("nItem")
		sProdcode= Request("sProdcode")
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
		sImagespec=Request("sImagespec")

		cStatus="9"
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty
chk_exict()
	if chkexict=1 then
		Response.redirect "PrintError.asp"
	end if
chk_product()
	If chkproduct=0 Then
		bError = True
		img_Prno = sErrImage
	End If
	If (Trim(nQuality) = "") Or Not IsNumeric(nQuality) Then
		bError = True
		img_nQuality = sErrImage
	End If

	If (Trim(nQuality) <> "") Or IsNumeric(nQuality) then
		if Cdbl(nQuality) <=0 then
			bError = True
			img_nQuality = sErrImage
		End if
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
	If chkvender=0 and svendcode<>"" Then
		bError = True
		img_sVendcode = sErrImage
	End If
	
	If bError = False Then
	  ' Create a recordset
	  
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode

		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  ppritem WHERE sPrno='" & sPrno &"'" &_
			"and nitem='" & nitem & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		'Add By Pik 2/06/2006 Error unique constraint
			sSQL = "delete from  ppritem WHERE sPrno='" & sPrno &"'  and nitem='" & nitem & "'"
			Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
'			Rs.Close
		' End Add
		  Rs.Open "ppritem", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		Rs("sPrno") = sPrno
		Rs("nItem") = nItem
		Rs("sProdcode") = sProdcode
		Rs("sProdname") = sProdname
		Rs("sVendcode") =sVendcode
		Rs("sProddesc") =sProddesc
		Rs("nQuality") =FormatNumber(nQuality,3)
		Rs("nPrice") =FormatNumber(nPrice, 3)
		Rs("nAmount") =FormatNumber((cDbl(FormatNumber(nQuality,3))*cDbl(FormatNumber(nPrice, 3))), 3)
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
if Request("approve")<>"app" then
     Response.Redirect "showpr.asp?sPrno=" & sPrno & "&Event=add"
else
     Response.Redirect "pprapp.asp?sPrno=" & sPrno & "&Event=add"
end if
'Response.wrie "Complete"
	ShowForm()
	else
		Show_Form()
	End If 
End Sub	

Sub SaveImage()
     dim objsmartupload,rsimagelib 	                      'for obj
     dim spath,fd,file,sql,smode,check1,check2            'other variable
     dim sprno,sfilename,nitem,sitemno   'for variable in table     
	  dim sImageType,sCatID,sCurcode,sUnitcode,sAcccode
     Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sdesc,nprice,sfile,dadddate,sempno
     set objsmartupload = server.createobject("aspSmartUpload.SmartUpload")'create Obj     
     objsmartupload.MaxFileSize = 2500000                                   'Max filesize 
     objsmartupload.Upload                                     'upload to objsmartupload

     ObjSmartUpload.DeniedFilesList = "bat,exe,com,asp,dll"    

     smode = objsmartupload.form("EVENT")
     'data from pre page      	
     if smode = "ADD" then
		sprno=objsmartupload.Form("sprno")
		nItem=objsmartupload.Form("nItem")
		sItemno=objsmartupload.Form("sItemno")
		sdesc = objsmartupload.Form("sdesc")
		sfile=objsmartupload.Form("sfile")
     else
		sprno=objsmartupload.Form("sprno")
		nItem=objsmartupload.Form("nItem")
		sItemno=objsmartupload.Form("sItemno")
		sdesc = objsmartupload.Form("sdesc")
		sfile=objsmartupload.Form("sfile")
     end if
   
	 spath = request.serverVariables("APPL_PHYSICAL_PATH")      'ª×èÍ ·Ò§physic ¢Í§ server
    spath = spath & "procurement\workflow\ppritemdesctImages"          'check path corp

     fd = objsmartupload.save(spath)           'save file to server at spath

	 for each file in objsmartupload.files
	 	 sfilename = file.filename             'ª×èÍ¢Í§ file ·Õèload ÁÒ
 	 next	
     'If not ObjSmartUpload.Files.IsMissing Then
     'sFileName = ObjSmartUpload.Files.FileName
     'End if                                   'samesame
 response.write sfilename
'	 check_blank()
'     check_filename() 

     set rsimagelib = server.createobject("ADODB.Recordset")

		 'update Table
         sql = "SELECT * from ppritem WHERE sprno='" & sprno & "'" &_
         " and nitem='" & nitem & "'" 
         rsimagelib.open sql,conn,adOpenKeySet,adLockPessimistic,adCmdtext
          If Not rsimagelib.EOF Then
		if sfilename<>"" then
			rsimagelib("simagespec") = sfilename
		end if
	     rsimagelib.update            
	     End IF

	 rsimagelib.close
	 set rsimagelib = nothing
'response.write "Add Complete"
     'response.redirect("../datamanager/pproductdesc.asp")
     'Response.Redirect "../datamanager/pproductdesc.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"&sCatid="&sCatID
End Sub
%>

<% 'check blank 
   'set objSmartUpload & Upload & set sMode

   Sub check_blank()
     dim str

     check1 = false
     check2 = false 
            
     if trim(sfilename)  = "" then 
        check1 = true
     end if
     if trim(stimgdesc) = "" then 
        check2 = true
     end if 

        
    if (sMode = "add") and (check1 or check2) then 
        str = "bbimagehome.asp?"         
        if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
        end if
        if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & "seimgdesc" & "&"    
        end if
        str = str & "event=" & sMode
        response.redirect(str)        
    end if

    if (sMode = "edit") and (check2) then 
        str = "sbimagehome.asp?"         
        if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
        end if
        if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & seimgdesc & "&"    
        end if
        str = str & "simgid=" & simgid & "&"
        str = str & "event=" & sMode
        response.redirect(str)        
    end if


   End Sub
%>

<%
   Sub check_filename()
    Dim Rs,sSql,str

    sSql = "SELECT * FROM pproduct where simage = '" & Ucase(sfilename) & "'"
    set Rs = Server.CreateObject("ADODB.Recordset")
    Rs.Open sSql,Conn,adOpenKeySet,adLockPessimistic,adCmdtext
   
	if not Rs.eof then
      str = "bbimagehome.asp?"
      if not check2 then
           str = str & "stimgdesc=" & stimgdesc & "&"    
      end if
      if (trim(seimgdesc) <> "") then
           str = str & "seimgdesc=" & seimgdesc & "&"    
      end if
      str = str & "simgid=" & simgid & "&"
      str = str & "event=" & sMode & "&"
      str = str & "Repeat=" & true      
      Response.Redirect(str)               	
   	end if

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
%><a href="../workflow/ppritemdescimages/<%=RS2("sfile")%>">
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
Sub chk_product()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pproduct Where sprodcode='" & sprodcode & "'" &_
		" and sExptype='" & Request("sExptype") & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		chkproduct=1
	else
		chkproduct=0
	End IF
End Sub
%>
<%
Sub chk_vender()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pvender Where sVendcode='" & sVendcode & "' and sename not like '(DO NOT USE)%'"

		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		chkvender=1
	else
		chkvender=0
	End IF
End Sub
%>
<%
Sub chk_Exict()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from ppritem Where sprno='" & sprno & "'" &_
		" and sprodcode='" & sprodcode & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		chkexict=1
	else
		chkexict=0
	End IF
End Sub
%>
