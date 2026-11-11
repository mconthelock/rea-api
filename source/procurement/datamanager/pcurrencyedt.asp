<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim scurcode,scurrency
	Dim img_currency,sbread
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
scurcode=Request("scurcode")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	If IsEmpty(Request.Form("EVENT")) Then
		If  scurcode <> "" Then
			sMode = "EDIT"
			Get_Currency(scurcode)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Currency(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Currency(scurcode)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM pcurrency WHERE scurcode='" & scurcode & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		scurcode = Rs("scurcode")
		scurrency = Rs("scurrency")
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
<title>Edit Currency :: Computer</title>
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
          <td align="right" valign="top" width="2%" height="2"><img src="../datamanager/images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"><b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Currency</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="95%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="23%"  align="Right"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Currency 
                          Code : </font></b></td>
                        <td width="77%"> <%If sMode="EDIT" then
Response.write sCurcode
Else%> 
                          <input type="text" name="scurcode" size="4" maxlength="3" value="<%=scurcode%>" <%=sbread%>>
                          <%=img_currency%> 
                          <%End if%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%"  align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Currency Code Data 
                          (Character) : Not more than 3 Characters ,Example : 
                          BTH.</font></td>
                      </tr>
                      <tr> 
                        <td width="23%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Currency 
                          : </font></b></td>
                        <td width="77%"> 
                          <input type="text" name="scurrency" size="50" maxlength="90" value="<%=scurrency%>">
                        </td>
                      </tr>
                      <tr>
                        <td width="23%"  align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Currency Unit Data(Character) 
                          : Not more than 30 Characters .</font></td>
                      </tr>
                      <tr>
                        <td width="23%"  align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC0000"><b>*</b></font> 
                          Required Field.</td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <%If sMode<>"ADD" then%> 
                    <input type=HIDDEN name=sCurcode value="<%=sCurcode%>">
                    <%End if%> 
                    <input type="submit" name="Submit" value="Save">
                    <input type="submit" name="Submit2" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../datamanager/datamanager.asp">Data Manager</a></font></p>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../datamanager/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../datamanager/images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
            <td> <br>
            <!--#include file="../inc/copyRight.inc"-->
		   </td>
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
Sub Update_Currency(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		scurcode = Request.Form("scurcode")
		scurrency = Request.Form("scurrency")

	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(scurcode) = "") Then
		bError = True
		img_currency = sErrImage
	End If


	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  Pcurrency WHERE scurcode='" & scurcode &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "Pcurrency", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("scurcode") = scurcode
		Rs("scurrency") = scurrency
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
     Response.Redirect "pcurrencyall.asp"
	else
		Show_Form()
	End If	' if Error 
End Sub	
%>