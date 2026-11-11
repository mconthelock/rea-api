<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim  sEmpNo,sEmpNo1
	Dim  img_sStep,sBread

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	If IsEmpty(Request.Form("EVENT")) Then
		sMode = "EDIT"
		Show_Form()
	Else
		Update_Flowstep(Request.Form("EVENT"))
	End If		
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
<title>Edit workflow Data</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2">&nbsp;</td>
    <td width="63%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="1" height="1"></td>
    <td width="36%" height="2" align="center"><img src="../images/spacer.gif" width="1" height="1"></td>
  </tr>
  <tr> 
    <td width="1%" height="2"><img src="../images/spacer.gif" width="1" height="1"></td>
    <td colspan="2" height="2" align="center" valign="top"> <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="datamanager/../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font><img src="../images/dot.gif"></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#6666CC"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Deputize Approve </font></b></font></td>
                <td width="8%" bgcolor="#3838AB" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#E6E6F7" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="95%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="50%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Approval
                          Employee no</b>. [OLD] </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                          </font></td>
                        <td width="50%"> 
                          <input type="text" name="sEmpno" size="12" maxlength="12" value="<%=sEmpNo%>">
                        </td>
                      </tr>
				     <tr> 
                        <td width="50%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b>Approval
                          Employee no</b>. [NEW] </font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
                          </font></td>
                        <td width="50%"> 
                          <input type="text" name="sEmpno1" size="12" maxlength="12" value="<%=sEmpNo1%>">
                        </td>
                      </tr>
                    </table>
                    <br>
					<input type="HIDDEN" name="EVENT" value='<%=sMode%>'>
                    <input type="submit" name="Submit" value="Save">
                    <input type="Reset" name="Submit" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="pappsteplist.asp">Other Approval Step</a></font></p>
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
Sub Update_Flowstep(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sEmpNo = Request.Form("sEmpNo")
		sEmpNo1 = Request.Form("sEmpNo1")

	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(SEMPNO) = "") Then
		bError = True
		img_sStep = sErrImage
	End If

	If (Trim(SEMPNO1) = "") Then
		bError = True
		img_sStep = sErrImage
	End If

	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "UPDATE PAPPstep Set SEMPNO = '"& SEMPNO1 &"' WHERE SEMPNO='" & SEMPNO &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText	
		End If
	  Set Rs = Nothing
      Response.Redirect "pappsteplist.asp"
	else
		Show_Form()
	End If	' if Error 
End Sub	

%>
