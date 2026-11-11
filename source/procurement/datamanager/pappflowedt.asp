<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" --><%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sFlow,sLevel,sFlowdesc
	Dim img_sFlow,sBread
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
sFlow=Request("sFlow")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	If IsEmpty(Request.Form("EVENT")) Then
		If  sFlow <> "" Then
			sMode = "EDIT"
			Get_Appflow(sFlow)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Appflow(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Appflow(sFlow)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM pAppflow WHERE sFlow='" & sFlow & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sFlow = Rs("sFlow")
		sLevel = Rs("sLevel")
		sFlowdesc=RS("sFlowdesc")
		sBread="ReadOnly"
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
<title><%=sMode%> ORG Code</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function Form_Load(){
	document.frmName.sFlow.focus();
}
-->
</SCRIPT>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="Form_Load()">
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
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><FONT COLOR="#FFCC66">..</FONT></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553">
          <td width="2%">&nbsp;</td>
          <td width="96%" valign="top"> <b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Approval 
                  Method :</font></b></font> <font color="#FFFFFF" size="2"><%=sMode%> 
                  ORG Code</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>" name="frmName">
                    <table width="95%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="23%" align="Right"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Flow 
                          No. : </font></b></td>
                        <td width="77%"> <%IF sMode="EDIT" then
			Response.write sFlow
		Else%> 
                          <input type="text" name="sFlow" size="12" maxlength="12" value="<%=sFlow%>" <%=sbread%>>
                          <%=img_sflow%> 
                          <%End if%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right" height="31">&nbsp;</td>
                        <td width="77%" height="31"><font color="#CC3333">Define 
                          Flow No. Data (Character) :Not more than 3 Characters 
                          ,Example : 311.</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right" height="31"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Level 
                          : </font></b></td>
                        <td width="77%" height="31"> 
                          <input type="text" name="sLevel" size="50" maxlength="90" value="<%=sLevel%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Define Organization 
                          Data For Approve Step .Order by Division, Department, 
                          Section (Character) :Not more than 60 Characters ,Example 
                          : 30101,30301,30302.</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Flow 
                          Description : </font></b></td>
                        <td width="77%"> 
                          <input type="text" name="sFlowdesc" size="50" maxlength="90" value="<%=sFlowdesc%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Description of Flow</font> 
                          <font color="#CC3333">(Character) :Not more than 90 
                          Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC0000"><b>*</b></font> 
                          Required Field.</td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <%If sMode<>"ADD"  and bError =false  then%> 
                    <input type=HIDDEN name=sFlow value="<%=sFlow%>">
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
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">Back</a> 
              | <a href="../datamanager/datamanager.asp">Data Manager</a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"></font><br>
            </p>
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
Sub Update_Appflow(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sFlow = Request.Form("sFlow")
		sLevel = Request.Form("sLevel")
		sFlowdesc = Request.Form("sFlowdesc")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(sFlow) = "") Then
		bError = True
		img_sFlow = sErrImage
	End If


	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  Pappflow WHERE sflow='" & sFlow &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "Pappflow", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("sFlow") = sFlow
		Rs("sLevel") = sLevel
		Rs("sFlowdesc") = sFlowdesc
	  Rs.Update	

	  Rs.Close
	  Set Rs = Nothing
     Response.Redirect "pappflowall.asp"
	else
		Show_Form()
	End If	' if Error 
End Sub	
%>