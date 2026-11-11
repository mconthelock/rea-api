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
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim spono,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Pono,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode
	Dim sUnitcode
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
                          No </font></b></td>
                        <td width="76%"> <%=sPono%></td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product Code</font></b></td>
                        <td width="76%">
						<%=sProdcode%>

                        </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product Name :</font></b></td>
                        <td width="76%"> 
                          <%=sProdname%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product Description :</font></b></td>
                        <td width="76%"> 
                          <%=sProddesc%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quantity :</font></b></td>
                        <td width="76%">
                          <%GetNumber(nQuality)%>&nbsp;<%Show_Unit(sUnitcode)%> 
						</td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UPrice :</font></b></td>
                        <td width="76%"> 
                          <%GetNumber(nPrice)%>&nbsp;<%Show_Currency(sCurcode)%>
						</td>
                      </tr>
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Amount :</font></b></td>
                        <td width="76%"> 
                          <%GetNumber(nAmount)%>&nbsp;<%Show_Currency(sCurcode)%>
                        </td>
                      </tr>
<!--                       <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UseFor 
                          </font></b></td>
                        <td width="76%"> 
                          <%=sUseFor%>
                        </td>
                      </tr> -->
                      <tr> 
                        <td width="24%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Comment</font></b></td>
                        <td width="76%"> 
                          <%=sComment%>
                        </td>
                      </tr>
                    </table>
                    <br>

                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2"> <A HREF="Javascript:window.close()"> Close </A>&nbsp; </td>
              </tr>
            </table>
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
		sCurcode= Request.Form("sCurcode")
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

		Rs("sPono") = sPono
		Rs("sProdcode") = sProdcode
		Rs("sProdname") = sProdname
		Rs("sProddesc") =sProddesc
		Rs("nQuality") =nQuality
		Rs("nPrice") =nPrice
		Rs("nAmount") =nAmount
		Rs("sCurcode") =sCurcode
		Rs("sComment") =sComment        
	  Rs.Update	

	  Rs.Close
	  Set Rs = Nothing
     Response.Redirect "../workflow/ppo.asp?sPono=" & sPono
	else
		Show_Form()
	End If	' if Error 
End Sub	
%>