<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!--#include file="../connas400.asp"-->
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
Response.Charset = "windows-874"
Session.CodePage = 874
%>
<%
'on error resume next
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,spono
	Dim sprno,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Prno,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode,img_cStatus
	Dim sUnitcode,skeyby,sRCno,nqualityo,nMaxQuality , spRecno
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		nitem=Request("nitem")
		sprno=Request("sprno")
		skeyby = Request("skeyby")
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
		sMode=Request("EVENT")
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

  	sSQL = "SELECT * FROM planreceiveitem WHERE nitem='" & nitem & "'" &_
	"and ID='" & srcno & "'"
    Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		spRecno=RS("ID")
		nItem = Rs("nItem")
		sProdcode=RS("sProdcode")
		sProdname=RS("sProdname")
		nQuality=RS("nQuality")
		nqualityo=nQuality
		sUnitcode=RS("sUnitcode")
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
<title><%=sMode%> Plan Receive Item</title>
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
                  Plan Receive Item</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center" height="26"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receive 
                          No. : </font></b></td>
                        <td width="75%"> <%=spRecno%></td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                          No. : </font></b></td>
                        <td width="75%"> <%=sPrno%></td>
                      </tr>
<!--                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Item 
                          No. : </font></b></td>
                        <td width="81%"><%=nItem%></td>
                      </tr> -->
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
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quantity 
                          : </font></b></td>
                        <td width="81%"> 
                          <input type="text" name="nQuality" size="19" maxlength="19" value="<%=nQuality%>">
                          <%=img_nQuality%> <%Show_Unit(sUnitcode)%> </td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name="sprno" value="<%=sprno%>">
                    <input type=HIDDEN name="nItem" value="<%=nItem%>">
                    <input type=HIDDEN name="sprodcode" value="<%=sprodcode%>">
                    <input type=HIDDEN name="sprodname" value="<%=sprodname%>">
                    <input type=HIDDEN name="nqualityo" value="<%=nqualityo%>">
                   <input type=HIDDEN name="sCurcode" value="<%=sCurcode%>">
                    <input type=HIDDEN name="sUnitcode" value="<%=sUnitcode%>"> 
                    <input type=HIDDEN name="sUsefor" value="<%=sUsefor%>"> 
                    <input type=HIDDEN name="sVendcode" value="<%=sVendcode%>"> 
                    <input type=HIDDEN name="skeyby" value="<%=Request("skeyby")%>">
                    <input type=HIDDEN name="spReCno" value="<%=spReCno%>">
                    <input type=HIDDEN name="spono" value="<%=Request("spono")%>">
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <input type=HIDDEN name="nPrice" value="<%=nPrice%>">
                    <input type=HIDDEN name=sExptype value="<%=sExptype%>">
                    <input type="submit" name="Submit" value="Plan Receive">
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
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sPrno = Request("sPrno")
		spRecno = Request("spRecno")
		nItem = Request("nItem")
		sProdcode= Request("sProdcode")
		sProdname= Request("sProdname")
		nQuality= Request("nQuality")
		nQualityo= Request("nQualityo")
		If (Trim(nQualityo) = "") Then
			nQualityo = 0
		end if
		sUnitcode= Request("sUnitcode")
		sVendcode=Request("sVendcode")
		
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

'	If (Trim(sPrno) = "") Then
'		bError = True
'		img_Prno = sErrImage
'	End If


	If (Trim(nQuality) = "") Or Not IsNumeric(nQuality) then 'Or Cdbl(nQuality) < 0 then 'Comment when implement SCM: make receiving transfer data program
		bError = True
		img_nQuality = sErrImage
	elseif (CDbl(nQuality)) > Sumitem2(sprno,sprodcode) Then
		bError = True
		img_nQuality = sErrImage
	End If
'	If (Trim(nAmount) = "") Then
'		bError = True
'		img_nAmount = sErrImage
'	End If
'	If (Trim(sVendcode) = "") Then
'		bError = True
'		img_sVendcode = sErrImage
'	End If

	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode

		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  planReceiveitem WHERE ID='" & sprecno &"'" &_
			"and nitem='" & nitem & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "planReceiveitem", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If

		Rs("ID") = Request("sprecno")
		Rs("nItem") = nItem
		Rs("sProdcode") = sProdcode
		Rs("sProdname") = sProdname
		Rs("nQuality") =nQuality
		Rs("sUnitcode") =sUnitcode      
		RS("sVendcode")=sVendcode
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
'	chkRec1()
'	chkRec2()
	
     Response.Redirect "../workflow/ppreceivedesc3.asp?sprno=" & sprno & "&nitem=" & nitem &"&spRecno=" & Request("sprecno")
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
	Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	sSQL2="Select sum(nQuality) sumrec from vprecitem where spono='" & spono & "'" &_
	" and sprodcode='" & sprodcode & "'"
	Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

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
			Set Rs4=Nothing
	else
		sSQL2="Select * from ppoitem where spono='" & spono & "'" &_
		" and sprodcode='" & sprodcode & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cStatus")="1"
			Rs4.Update
			RS4.Close

		sSQL2="Select * from preceiveitem where sRecno='" & Request("sRCno") & "'" &_
		" and sprodcode='" & sprodcode & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cStatus")="1"
			Rs4.Update
			RS4.Close

			sSQL="Select * from ppr where sprno='" & sPrno & "'"
 			 Rs4.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			Do while Not RS4.EOF	
				RS4("cPo")="2"
				RS4.Update
			RS4.MoveNext
			Loop
			Rs4.Close
			Set Rs4=Nothing
	End if
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
	Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	if Rs3.EOF then
		sSQL2="Select * from ppo where spono='" & spono & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cstatus")="3"
			Rs4.Update
			RS4.Close
	else
		sSQL2="Select * from ppo where spono='" & spono & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cstatus")="2"
			Rs4.Update
			RS4.Close
	End if
	Rs3.Close

	sSQL2="Select * from ppo where ppo.sRefno='" & sprno & "'" &_
	" and cstatus<>'3'" &_
	" and cstatus<>'9'"
	Rs3.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	if RS3.EOF then
		sSQL2="Select * from ppr where sprno='" & sprno & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cpo")="3"
			Rs4.Update
			RS4.Close
	else
		sSQL2="Select * from ppr where sprno='" & sprno & "'"
		  Rs4.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
			RS4("cpo")="2"
			Rs4.Update
			RS4.Close
	End if
			Set Rs4=Nothing

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
			sumitem=cDbl(nQuality)-cDbl(RS2("sumitems"))
		else
			Sumitem=nQuality
		end if
	End if
	Rs2.Close
	Set Rs2=Nothing
End Function

Function SumItem2(sprno,sprodcode)
	Dim Rs2,Rs3,RS4,sSQL2
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")
	sSQL2="Select Sum(nQuality) sumitems from vprecitem where sprno='" & sprno &_
	"' and sprodcode='" & sProdcode	& "'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		
	sSQL2="Select nQuality from ppritem where sprno='" & sprno &_
	"' and sprodcode='" & sProdcode	& "'"
	  Rs3.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText

	If Not RS2.EOF then
		if RS2("Sumitems")<>"" then
			sumitem2=cDbl(RS3("nQuality"))-cDbl(RS2("sumitems"))
		else
			Sumitem2=cDbl(RS3("nQuality"))

		end if
	End if
	Rs2.Close
	Set Rs2=Nothing
	Rs3.Close
	Set Rs3=Nothing
End Function
%>