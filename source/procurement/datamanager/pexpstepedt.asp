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
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim nstep,sexptype,sstep,sempno,sdesc,sExpDESC
	Dim img_nstep,sbread,img_desc
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
nstep=Request("nstep")
sexptype=Request("sexptype")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	If IsEmpty(Request.Form("EVENT")) Then
		If  nstep <> "" Then
			sMode = "EDIT"
			Get_ExpStep(nstep)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_ExpStep(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_ExpStep(nstep)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM pexpstep WHERE nstep='" & nstep & "'" &_
"and sexptype='" & sexptype & "'"
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		nstep = RS("nstep")
		sexptype = Rs("sexptype")
		sExptype = Rs("sexptype")	
		sstep = RS("sstep")
		sempno = RS("sempno")
		sdesc = Rs("sdesc")
		sbread = "ReadOnly"
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
<title><%=sMode%>  Expense Type Approval Step </title>
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
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><FONT COLOR="#FFCC66">..</FONT></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../security/images/y02.jpg" width="12" height="12"></td>
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Approval 
                  Method : <%=sMode%> 
                  Expense Type Approval Step</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="23%" align="Right"><b><font color="#CC0000"><b>* 
                          </b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Step 
                          No. : </font></b></td>
                        <td width="77%"> <%If sMode="EDIT" then
			Response.write nStep
		Else%> 
                          <input type="text" name="NSTEP" size="3" maxlength="2" value="<%=nstep%>" <%=sbread%>>
                          <%Response.write img_nstep
End if%> </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Define Step for 
                          Approval(Numeric) : Not more than 99 .Example : 5</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right"><b><font color="#CC0000"><b> 
                          *</b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Expense Type : </font></b></td>
                        <td width="77%"> <% if sMode="ADD" then %> 
                          <SELECT NAME="SEXPTYPE" SIZE=1 <%=sBread%>>
                            <% List_Pexptype(sexptype) %> 
                          </SELECT>
                          <%else
Show_Pexptype(sexptype)%> <!--                          <input type="text" name="SEXPTYPE" size="50" maxlength="50" value='<%pExptype_Desc(sExptype)%>' <%=sbread%>> --> 
                          <% END IF%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Expense Type Data. 
                          Select from List.</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Step 
                          : </font></b></td>
                        <td width="77%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          <select name="SSTEP">
                            <% List_Pflowstep(sStep) %> 
                          </select>
                          </font> </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Step of Complete 
                          Flow Data. Select from List.</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Employee 
                          No : </font></b></td>
                        <td width="77%"> 
                          <input type="text" name="SEMPNO" size="12" maxlength="12" value="<%=SEMPNO%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right" valign="top">&nbsp;</td>
                        <td width="77%"><font color="#CC3333">Employee to Approve 
                          this Step (Character,Employee No.) : Not more than 12 
                          Characters.</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Description 
                          : </font></b></td>
                        <td width="77%"> 
                          <textarea name="SDESC" cols="60"><%=sdesc%></textarea><%=img_desc%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right" valign="top">&nbsp;</td>
                        <td width="77%"><font color="#CC3333"> Description(Character) 
                          : Not more than 120 Characters.</font></td>
                      </tr>
                      <tr> 
                        <td width="23%" align="Right" valign="top">&nbsp;</td>
                        <td width="77%"><font color="#CC0000"><b>*</b></font> 
                          Required Field.</td>
                      </tr>
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <%if sMode <> "ADD" then%> 
                    <input type=HIDDEN name=sExptype value="<%=sExptype%>">
                    <input type=HIDDEN name=nStep value="<%=nStep%>">
      <!--              <input type=HIDDEN name=sStep value="<%=sStep%>"> -->
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
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../security/images/y04.jpg" width="12" height="12"></td>
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
</body>
</html>

<% End Sub %>

<%
Sub Update_ExpStep(sUpdateMode)
	Dim sErrImage, Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		nstep = Request.Form("nstep")
		sexptype = Request.Form("sexptype")
'		sExptype = Session("sExptype")
		sStep= Request.Form("sstep")
		sEmpno= Request.Form("sempno")
		sDesc= Request.Form("sdesc")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(nStep) = "") Then
		bError = True
		img_nstep = sErrImage
	End If
	If Len(sDesc) >120 Then
		bError = True
		img_desc = sErrImage
	End If
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  pExpstep WHERE nstep = " & nStep & " and sExptype = '"&sExptype & "'"
'		RS.Open sSQL,Conn,1,3
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
			sSQL = "SELECT * FROM  pExpstep"
'		RS.Open sSQL,Conn,1,3
'		   Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdTable
	  Rs.Open "Pexpstep", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		   Rs.AddNew
		End If
		'Response.Write		mid(nStep,1,2)
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		
		Rs("nStep") = mid(nStep,1,2)
		Rs("sExptype") = mid(sExptype,1,2)
		Rs("sStep") = sStep
		Rs("sEmpno") = sEmpno
		Rs("sDesc") = sDesc
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
		Session("sExptype") = ""		
     Response.Redirect "pexpstepall.asp"
	else
response.write smode
		Show_Form()
	End If	' if Error 
End Sub	
%>