<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../List.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim spono,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Pono,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode
	Dim sUnitcode,sprno
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		sProdcode=Request("sProdcode")
		spono=Request("spono")
	If IsEmpty(Request.Form("EVENT")) Then
		If  spono <> "" and sprodcode <> "" Then
			sMode = "EDIT"
			Get_Poitem(sPono)			
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Poitem(Request.Form("EVENT"))
	End If	
	
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Poitem(sPono)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  	sSQL = "SELECT * FROM ppoitem WHERE sPono='" & sPono & "'" &_
	"and sProdcode='" & sProdcode & "'"
'	response.write sSQL
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sPono=RS("sPono")
		sProdcode=RS("sProdcode")
		sProdname=RS("sProdname")
		sProddesc=RS("sProddesc")
		nQuality=RS("nQuality")
		nPrice=RS("nPrice")
		nAmount=RS("nAmount")
		sCurcode=RS("sCurcode")
		sUnitcode=RS("sUnitcode")
		sComment=RS("sComment")
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
<title>Edit Vendors :: Computer</title>
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
                  POItem</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O 
                          No :</font></b></td>
                        <td width="76%"> <%=sPono%></td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product code :</font></b></td>
                        <td width="76%">
					<%if sMode="EDIT" then%>
						<%=sProdcode%>

					<%Else%>
                          <input type="text" name="sprodcode" size="13" maxlength="12" value="<%=sprodcode%>">
					<%End if%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product name :</font></b></td>
                        <td width="76%"> <%'=sProdname%>
                           <input type="text" name="sProdname" size="100" value="<%=sProdname%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right" valign = "top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product Description :</font></b></td>
                        <td width="76%"> <%'=sProddesc%>
                           <textarea name="sProddesc" cols="45" rows="3"><%=sProddesc%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quantity :</font></b></td>
                        <td width="76%"> 
                          <input type="text" name="nQuality" size="19" maxlength="19" value="<%=nQuality%>">
                          <%=img_nQuality%> <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b> </b></font>
							<%Show_Unit(sUnitcode)%> 
						</td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UPrice :</font></b></td>
                        <td width="76%"> 
                          <input type="text" name="nPrice" size="19" maxlength="19" value="<%=nPrice%>"><%Show_Currency(sCurcode)%>
                          <%=img_nPrice%> </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Amount :</font></b></td>
                        <td width="76%"> <%=nAmount%>
                           </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right" valign=top><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Comment :</font></b></td>
                        <td width="76%">
                           <textarea name="sComment" cols="45" rows="3"><%=sComment%></textarea>
                        </td>
                      </tr>					  
<!--                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Currency :</font></b></td>
                        <td width="76%"> 
                           <SELECT NAME="sCurrency" SIZE=1>
                            <%List_Currency(sCurcode)%> 
                          </SELECT> -->
                        </td>
                      </tr>
<!--                       <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UseFor 
                          </font></b></td>
                        <td width="76%"> 
                          <input type="text" name="sUseFor" size="13" maxlength="120" value="<%=sUseFor%>">
                        </td>
                      </tr> 
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Comment :</font></b></td>
                        <td width="76%"> <%=sComment%>
                           <textarea name="sComment" cols="45" rows="3"><%=sComment%></textarea>
                        </td>
                      </tr>-->
                    </table>
                    <br>

                    <input type=HIDDEN name="sPono" value="<%=sPono%>">
                    <input type=HIDDEN name="sProdcode" value="<%=sProdcode%>">
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
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
Sub Update_Poitem(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sPono = Request("sPono")
		sProdcode= Request("sProdcode")
		sProdname= Request.Form("sProdname")
		sProddesc= Request.Form("sProddesc")
		nQuality= Request.Form("nQuality")
		nPrice= Request.Form("nPrice")
		nAmount= Request.Form("nAmount")
		sUnitcode= Request("sUnit")
		sComment= Request.Form("sComment")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(sPono) = "") Then
		bError = True
		img_Pono = sErrImage
	End If
	If (Trim(nQuality) = "") Or Not IsNumeric(nQuality) Then
		bError = True
		img_nQuality = sErrImage
	End If
	If (Trim(nPrice) = "") Then
		bError = True
		img_nPrice = sErrImage
	End If
'	If (Trim(nAmount) = "") Then
'		bError = True
'		img_nAmount = sErrImage
'	End If

	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  ppoitem WHERE sPono='" & sPono &"'" &_
			"and sProdcode='" & sprodcode & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "ppoitem", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If

'		Rs("sPono") = sPono
'		Rs("sProdcode") = sProdcode
'Remove Comment 	13/12/2010
		Rs("sProdname") = sProdname
		Rs("sProddesc") =sProddesc
		Rs("nQuality") =nQuality
		Rs("nPrice") =nPrice
		Rs("nAmount") =formatnumber(nQuality, 3)*formatnumber(nPrice, 4)
'		Rs("sUnitcode") =sUnitcode
'		Rs("sCurcode") =Request("sCurrency")
'		Rs("sComment") =sComment        
		Rs("sComment") =sComment        
	  Rs.Update	

	  Rs.Close
		sSQL="Select srefno from ppo where spono='" & spono & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText	 
			if Not Rs.EOF then
				sPrno=Rs("sRefno")
			End if 
		  Rs.Close
'		sSQL="Select * from ppritem where sprno='" & sprno & "'"
'		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText	 
'			IF Not RS.EOF then
'				Rs("sProddesc") =sProddesc
'				Rs("nQuality") =nQuality
'				Rs("nPrice") =nPrice
'				Rs("nAmount") =formatnumber(nQuality)*formatnumber(nPrice)
'			  Rs.Update	
'			End if
	  Set Rs = Nothing
		if Request("chk") = "1" then
			Response.Redirect "../workflow/ppo.asp?sPono=" & sPono
		else
			Response.Redirect "../workflow/ppoprintedt.asp?sPono=" & sPono
		end if
	else
		Show_Form()
	End If	' if Error 
End Sub	
%>