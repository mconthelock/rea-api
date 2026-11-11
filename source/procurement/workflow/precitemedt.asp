<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../Connas400.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<!--#include file="ppo.inc"-->
<!--#include file="scm.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
'on error resume next
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,spono
	Dim sprno,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Prno,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode,img_cStatus
	Dim sUnitcode,sRecno,sRCno,nqualityo,nMaxQuality,page,maxpage
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		nitem=Request("nitem")
		sprno=Request("sprno")
		sRecno = Request("sRecno")
		sRcno = Request("sRcno")
		spono = Request("spono")
		sProdcode = Request("sProdcode")
		nMaxQuality=Request("nMaxQuality")
  
	If IsEmpty(Request.Form("EVENT")) Then
		If  nitem<> "" Then
			sMode = "EDIT"
			Get_Pritem(sPrno)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		sMode=Request.Form("EVENT")
		Update_Pritem(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Pritem(sPrno)
	Dim sSQL, Rs,RS1
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
   Set Rs = Server.CreateObject("ADODB.RecordSet")

  	sSQL = "SELECT * FROM preceiveitem WHERE nitem='" & nitem & "'" &_
	" and srecno='" & srcno & "'"
    Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
'	response.write sSQL
'	response.end
   
	If Not Rs.EOF Then
		 
		sRecno=RS("sRecno")
		nItem = Rs("nItem")
		sProdcode=RS("sProdcode")
		sProdname=RS("sProdname")
		nQuality=RS("nQuality")
		nqualityo=nQuality
		nPrice=RS("nPrice")
		nAmount=RS("nAmount")
		sCurcode=RS("sCurcode")
		sUseFor=RS("sUseFor")
		sComment=RS("sComment")
		sUnitcode=RS("sUnitcode")
		cStatus=RS("cStatus")
		sVendcode=RS("sVendcode")
		sbread="ReadOnly"
	Else
		
		nqualityo=0
  		sSQL = "SELECT * FROM vppoitem WHERE  spono='" & spono & "'" &_
		" and sprodcode='" & sprodcode & "'"
		 Set Rs1 = Server.CreateObject("ADODB.RecordSet")
		 Rs1.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	     If Not Rs1.EOF Then
		    
			sPono=RS1("sPono")
'			nItem = Rs1("nItem")
			sProdcode=RS1("sProdcode")
			sProdname=RS1("sProdname")
			nMaxQuality=RS1("nQuality")
			nQuality=Sumitem(RS1("nQuality"),sprno,sprodcode)
			nPrice=RS1("nPrice")
			nAmount=RS1("nAmount")
			sCurcode=RS1("sCurcode")
'			sUseFor=RS1("sUseFor")
			sComment=RS1("sComment")
			sUnitcode=RS1("sUnitcode")
'			cStatus=RS1("cStatus")
			sVendcode=RS1("sVendcode")
			sbread="ReadOnly"
		end if
		sMode = "ADD"
		rs1.close
		set rs1=nothing
	End If
	rs.close
	set rs=nothing
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
                <td width="92%" bgcolor="#DB4B37" height="26"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%></font></b></font><font color="#FFFFFF" size="2"> 
                  Receive Item</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center" height="26"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Receive 
                          No. : </font></b></td>
                        <td width="75%"> <%=sRecno%></td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O 
                          No. : </font></b></td>
                        <td width="75%"> <%=sPono%></td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Item 
                          No. : </font></b></td>
                        <td width="81%"><%=nItem%></td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Code : </font></b></td>
                        <td width="81%"> <%=sProdcode%> </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Name : </font></b></td>
                        <td width="81%"> <%=sProdname%> </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quality 
                          : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="nQuality" size="19" maxlength="19" value="<%=nQuality%>">
                          <%=img_nQuality%> <%Show_Unit(sUnitcode)%> </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Price/Unit 
                          : </font></b></td>
                        <td width="81%"> <%=nPrice%>
<!--                          <input type="text" name="nPrice" size="19" maxlength="19" value="<%=nPrice%>"> -->
                           <%Show_Currency(sCurcode)%> </td>
                      </tr>
                      <!--                      <tr> 
                        <td width="19%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Amount</font></td>
                        <td width="81%"> 
                          <input type="text" name="nAmount" size="19" maxlength="19" value="<%=nAmount%>"><%=img_nAmount%>
                        </td>
                      </tr>  
                      <tr> 
                        <td width="19%" align="Right" height="33"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UseFor 
                          : </font></b></td>
                        <td width="81%" height="33"> 
                          <%List_Section(sUsefor)%>
                        </td>
                      </tr>
                      <%'if sMode<>"ADD" then%>                      <tr> 
                        <td width="19%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Status</font></td>
                        <td width="81%"> 
                          <input type="text" name="cStatus" size="2" maxlength="1" value="<%=cStatus%>"><%=img_cStatus%>
                        </td>
                      </tr> --> <%'end if%> 
                    </table>
                    <br>
                    <input type=HIDDEN name="sprno" value="<%=sprno%>">
                    <input type=HIDDEN name="nItem" value="<%=nItem%>">
                    <input type=HIDDEN name="sprodcode" value="<%=sprodcode%>">
                    <input type=HIDDEN name="sprodname" value="<%=server.htmlEncode(sprodname)%>">
                    <input type=HIDDEN name="nqualityo" value="<%=nqualityo%>">
                    <input type=HIDDEN name="nMaxquality" value="<%=nMaxquality%>">
                   <input type=HIDDEN name="sCurcode" value="<%=sCurcode%>">
                    <input type=HIDDEN name="sUnitcode" value="<%=sUnitcode%>"> 
                    <input type=HIDDEN name="sUsefor" value="<%=sUsefor%>"> 
                    <input type=HIDDEN name="sComment" value="<%=sComment%>"> 
                    <input type=HIDDEN name="sVendcode" value="<%=sVendcode%>"> 
                    <input type=HIDDEN name="sRCno" value="<%=Request("sRCno")%>">
                    <input type=HIDDEN name="sReCno" value="<%=Request("sReCno")%>">
                    <input type=HIDDEN name="spono" value="<%=Request("spono")%>">
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name="nPrice" value="<%=nPrice%>">
                    <input type=HIDDEN name="pageno" value="<%=Request("pageno")%>">
                    <input type=HIDDEN name="maxpage" value="<%=Request("maxpage")%>">
                    <input type=HIDDEN name=sExptype value="<%=sExptype%>">
					<input type="submit" name="Submit" value="Receive">
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
'response.write "1[" & Request("sPrno") & "]<br>"
'response.write "2[" &  Request("sRecno") & "]<br>"
'response.write "3[" &  Request("spono") & "]<br>"
'response.write "4[" &  Request("nItem") & "]<br>"
'response.write "5[" &  Request("sProdcode") & "]<br>"
'response.write "7[" &  Request("nQuality") & "]<br>"
'response.write "8[" &  Request("nQualityo") & "]<br>"
'response.write "9[" &  Request("nMaxQuality") & "]<br>"

'response.write "10[" &  Request("nPrice") & "]<br>"
'response.write "11[" &  Request("sComment") & "]<br>"
'response.write "12[" &  Request("sCurcode") & "]<br>"
'response.write "13[" &  Request("sUseFor") & "]<br>"
'response.write "14[" &  Request("sUnitcode") & "]<br>"
'response.write "15[" &  Request("sVendcode") & "]<br>"
'response.write "16[" &  Request("sRCno") & "]<br>"
'response.end

	' Check to see if Start date is a valid date
		sPrno = Request("sPrno")
		sRecno = Request("sRecno")
		spono = Request("spono")
		nItem = Request("nItem")
		sProdcode= Request("sProdcode")
		sProdname= Request("sProdname")
		nQuality= Request("nQuality")
		nQualityo= Request("nQualityo")
		nMaxQuality= Request("nMaxQuality")
		If (Trim(nQualityo) = "") Then
			nQualityo = 0
		end if
		nPrice= Request("nPrice")
		sComment= Request("sComment")
		sCurcode= Request("sCurcode")
		sUseFor= Request("sUseFor")
		sUnitcode= Request("sUnitcode")
		sVendcode=Request("sVendcode")
		cStatus="9"

	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

'	If (Trim(sPrno) = "") Then
'		bError = True
'		img_Prno = sErrImage
'	End If
'response.write Sumitem(nQuality,sprno,sprodcode)
	
 'response.write "[if else]"&Sumitem2(spono,sprodcode)
  ' Sumitem2(spono,sprodcode)
  ' response.end
	If (Trim(nQuality) = "") Or Not IsNumeric(nQuality) then 'Or Cdbl(nQuality) <= 0 then 'Comment when implement SCM: make receiving transfer data program
		bError = True
		img_nQuality = sErrImage
'	elseif (CDbl(nQuality) - CDbl(nQualityo)) > Sumitem2(spono,sprodcode) Then
	elseif (CSng(nQuality) - CSng(nQualityo)) > Sumitem2(spono,sprodcode) Then
		bError = True
		img_nQuality = sErrImage
	End If
		
'	If (Trim(nPrice) = "") Then
'		bError = True
'		img_nPrice = sErrImage
'	End If

	If cDbl(nQuality) > cDbl(nMaxQuality)  Then
		bError = True
		img_nQuality = sErrImage
	End If
'	If (Trim(sVendcode) = "") Then
'		bError = True
'		img_sVendcode = sErrImage
'	End If


	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
      
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  pReceiveitem WHERE sRecno='" & sRCno &"'" &_
			"and nitem='" & nitem & "'"
'			sSQL="update preceiveitem set (srecno='" & Request("sRCno") &"'" _
'			& " ,nitem='" & nItem & "' ,sprodcode= '" & sProdcode & "', sprodname='" & sProdname & "'" _
'			& " ,nquality='" & nQuality & "' ,nprice='" & FormatNumber(nPrice, 3) & "'" _
'			& " ,namount='" & FormatNumber((cDbl(nQuality)*cDbl(nPrice)), 3) & "',scurcode='" & sCurcode & "'" _
'			& " ,cstatus='1' ,susefor='" & sUseFor & "' ,sunitcode='" & sUnitcode & "' ,svendcode='" & sVendcode & "'" _ 
'			& " scomment='" & sComment & "' where sRecno='" & sRCno &"' and nitem='" & nitem & "'"

		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "pReceiveitem", conn, adOpenKeySet, adLockPessimistic, adCmdTable
'			sSQL = "insert into preceiveitem(srecno, nitem, sprodcode, sprodname, nquality, nprice, namount,scurcode, cstatus, susefor, sunitcode, svendcode, scomment)" _
'			& " values('" & Request("sRCno") &"','" & nItem & "','" & sProdcode & "','" & sProdname & "','" & nQuality & "','" & FormatNumber(nPrice, 3) & "'" _
'			& ",'" & FormatNumber((cDbl(nQuality)*cDbl(nPrice)), 3) & "','" & sCurcode & "','1','" & sUseFor & "','" & sUnitcode & "','" & sVendcode & "' ,'" & sComment & "')"
'			sSQL = "insert into preceiveitem(srecno, nitem, sprodcode, sprodname, nquality, nprice, namount,scurcode, cstatus, susefor, sunitcode, svendcode, scomment)" _
'			& " values('" & Request("sRCno") &"','" & nItem & "','" & sProdcode & "','" & sProdname & "','" & nQuality & "','" & nPrice & "'" _
'			& ",'" & (cDbl(nQuality)*cDbl(nPrice)) & "','" & sCurcode & "','1','" & sUseFor & "','" & sUnitcode & "','" & sVendcode & "' ,'" & sComment & "')"
			
'			response.write sSQL
'			response.end
'			Rs.close
'		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		  Rs.AddNew
		End If
		
		Rs("sRecno") = Request("sRCno")
		Rs("nItem") = nItem
		Rs("sProdcode") = sProdcode
		Rs("sProdname") = sProdname
		Rs("nQuality") =nQuality
		Rs("nPrice") =FormatNumber(nPrice, 3)
		Rs("nAmount") =FormatNumber((cDbl(nQuality)*cDbl(nPrice)), 3)
		Rs("nPrice") =nPrice
		Rs("nAmount") =(cDbl(nQuality)*cDbl(nPrice))
		Rs("sCurcode") =sCurcode
		Rs("sUseFor") =sUseFor
		Rs("sComment") =sComment
		Rs("sUnitcode") =sUnitcode
		Rs("cStatus") ="1"        
		RS("sVendcode")=sVendcode
'		response.write Request("sRCno") & "[" & nItem & "][" & sProdname & "]["	& nQuality & "]<br>"
'		response.write FormatNumber(nPrice, 3) & "[" & FormatNumber((cDbl(nQuality)*cDbl(nPrice)), 3) & "]<br>"
'		response.write sCurcode & "[" & sUseFor & "][" & sComment & "][" & sUnitcode & "]<br>"
'		response.write sVendcode

	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
	chkRec1()
	chkRec2()

	'########## SEND RECEIVING DATA TO SCM ##########
	if isSCMVender(sVendCode, conn) and timeIsOk() then
'		sendRecDataToSCM sPoNo, sRcNo, sRecNo, day(date) & "/" & month(date) & "/" & year(date), sProdCode, nQuality, conn
		sendRecDataToSCM sPoNo, sRcNo, sRecNo, day(FormatDateTime(Date, 2)) & "/" & month(FormatDateTime(Date, 2)) & "/" & year(FormatDateTime(Date, 2)), sProdCode, nQuality,nItem, conn
	end if
	'########## END SEND RECEIVING DATA TO SCM ##########
'	response.end
     Response.Redirect "../workflow/preceivedesc2.asp?sPrno=" & sPrno & "&nitem=" & nitem &"&sRecno=" & Request("sRCno")&"&pageno="& Request("pageno")&"&maxpage="&Request("maxpage")
'     Response.Redirect "../workflow/pprapp.asp?sPrno=" & sPrno & "&nitem=" & nitem &"&sRecno=" & Request("sRCno")
'Response.wrie "Complete"
	ShowForm(sprno)
	else
		Show_Form()
	End If 
End Sub	

Sub chkRec1()
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")
	Set Rs4 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select sum(nQuality) sumppr from ppoitem where spono='" & sPono & "'" &_
	" and sprodcode='" & sprodcode & "'"
'	response.write sSQL2
'	response.end
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	sSQL2="Select sum(nQuality) sumrec from vprecitem where spono='" & spono & "'" &_
	" and sprodcode='" & sprodcode & "'"
	Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
'	response.write Cdbl(RS2("sumppr")) & " -" & CDbl(RS3("sumrec")) 
'	if Cdbl(RS2("sumppr"))=CDbl(RS3("sumrec")) then
'		response.write "eq"
'	else
'		response.write "not eq"	
'	end if
'	response.end
if RS2("sumppr")<>"" or RS3("sumrec")<>"" then
	if Cdbl(RS2("sumppr"))=CDbl(RS3("sumrec")) then
		sSQL2="Select * from ppoitem where spono='" & spono & "'" &_
		" and sprodcode='" & sprodcode & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cStatus")="2"
			Rs4.Update
			RS4.Close
		sSQL2="Select * from preceiveitem where sRecno='" & Request("sRCno") & "'" &_
		" and sprodcode='" & sprodcode & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cStatus")="2"
			Rs4.Update
			RS4.Close
		sSQL2="Select * from ppo where spono='" & spono & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cstatus")="2"
			Rs4.Update
			RS4.Close
			Set Rs4=Nothing
	End if
end if
	RS2.Close
	RS3.Close
	Set Rs2=Nothing
	Set Rs3=Nothing
End Sub

Sub chkRec2()
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")
	Set Rs4 = Server.CreateObject("ADODB.RecordSet")

'	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	sSQL2="Select * from ppoitem where spono='" & spono & "'" &_
	" and cstatus<>'2'"
'	response.write sSQL2 & "<br>"
	Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	if Rs3.EOF then
		sSQL2="Select * from ppo where spono='" & spono & "'"
'			response.write sSQL2 & "<br>"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cstatus")="3"
			Rs4.Update
			RS4.Close
	End if
	Rs3.Close

	sSQL2="Select * from ppo where ppo.sRefno='" & sprno & "'" &_
	" and cstatus<>'3'" &_
	" and cstatus<>'9'"
'				response.write sSQL2 & "<br>"
	Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	if RS3.EOF then
		sSQL2="Select * from ppr where sprno='" & sprno & "'"
'					response.write sSQL2 & "<br>"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cpo")="3"
			Rs4.Update
			RS4.Close
			'Place update/add product quantity to IS Small Stock Control's script here

	End if
			Set Rs4=Nothing
'response.end
'	End if
'End if
'	RS2.Close
	RS3.Close
	Set Rs2=Nothing
	Set Rs3=Nothing
End Sub

Function SumItem(nQuality,sprno,sprodcode)
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select Sum(nQuality) sumitems from vprecitem where sprno='" & sprno &_
	"' and sprodcode='" & sProdcode	& "'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText

	If Not RS2.EOF then
		if RS2("Sumitems")<>"" then
			
			
			SumItem= (cDbl(nQuality)*1000 - cDbl(RS2("sumitems")) * 1000)/1000
			'response.write "xxx"&(36000 - 35935)/1000
			'response.write cDbl(RS2("sumitems"))
			
		else
			Sumitem=nQuality
		end if
	End if
	Rs2.Close
	Set Rs2=Nothing
End Function

Function SumItem2(spono,sprodcode)
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")
	
	sSQL2="Select Sum(nQuality) sumitems from vprecitem where spono='" & spono &_
	"' and sprodcode='" & sProdcode	& "'"
	'response.write sSQL2
	'response.end
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	
	sSQL2="Select nQuality from ppoitem where spono='" & spono &_
	"' and sprodcode='" & sProdcode	& "'"

	Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	'response.write Rs3("nQuality")

	  
	If Not Rs2.EOF then
		if Rs2("sumitems")<>"" then
			SumItem2=cDbl(Rs3("nQuality"))-cDbl(Rs2("sumitems"))
			' response.end
		else
			SumItem2=nQuality
			'response.write "eeee"
			 'response.end
		end if
	End if
	
	Rs2.Close
	Set Rs2=Nothing
	Rs3.Close
	Set Rs3=Nothing
End Function
%>