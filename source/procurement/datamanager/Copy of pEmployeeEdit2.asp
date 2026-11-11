<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../List.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: worakrit yodthonglert
	' Version: 1.01
%>		

<%

    Dim  Rs,sSQL,sStatus,sScriptPath,smode
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim scurcode,scurrency
	'Dim img_currency,sbread
	'Dim sScriptPath
	Dim img_sEmpno, img_sName, img_sSurname, img_sTName,  img_sTSurname, img_sLogin ,img_sPassword
	Dim img_sDivision, img_sDepartment, img_sSection, img_sPosition,img_sStatus,Img_sOldpassword,Img_RecMail
	Dim sDivision, sDepartment, sSection, sPosition
	Dim sDivcode, sDepcode, sSeccode, sPoscode,sStatuscode,sStatusDesc
	Dim sEmpno, sDivc, sDepc, sSecc, sPO, sName, sSurname, sTName, sTSurname
	Dim sLogin, sPassword, sPassword1,sOldpassword1,sOldpassword2,cRecMail,sRecMail
	Dim bError, sbread,status
'	Dim sMode
	Dim sErrImage
	bError = False
	sErrImage = "<IMG SRC='../images/icerror.gif'>"
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
sEmpno=Request("sEmpno")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")		
'	sUserid = Request("UID")
	If IsEmpty(Request.Form("EVENT")) Then
		If  sEmpno <> "" Then
			sMode = "EDIT"
			Get_Employee(sEmpno)
		Else
			sDivcode = "00"
			sDepcode = "00"
			sSeccode = "00"
			sPoscode = "00"
			sMode = "ADD"
		End If		
		Show_Form()
	Else
		Update_Employee(Request.Form("EVENT"))
	End If			
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Employee(sEmpno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM aEmployee WHERE sEmpno ='" & sEmpno & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sEmpno = Rs("sEmpno")
		sDivcode = Rs("sDivcode")
		sDepcode = Rs("sDepcode")
		sSeccode = Rs("sSeccode")
		sPoscode = Rs("sPoscode")
		sStatuscode = Rs("cStatus")
		sName = Rs("sName")
		sSurname = Rs("sSurname")
		sTName = Rs("sTName")
		sTSurname = Rs("sTSurname")
		sLogin = Rs("sLogin")
		sOldPassword1 = RS("sPassword")
		cRecMail=RS("cRecMail")
		sRecMail=RS("sRecMail")
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
<title><%=sMode%> Employee</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function Form_Load(){
	document.FormAddUser.sEmpno.focus();
}
-->
</SCRIPT>
</head>
<%	if Mid(Session("sAuthor"),19,3) = "AED" and sMode <> "EDIT" Then%>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="Form_Load()" >
<%
Else
	response.write "<body bgcolor='#FFFFFF' leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
End IF
%>
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
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"><b><img src="../images/plus.gif" width="9" height="9">  
            USER MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553">  
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top">  
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>  
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><%=sMode%>  
                  Employee </b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top">  
                <td colspan="2">  
                  <form method="post" action='<%=sScriptPath%>' name="FormAddUser">
                    <table width="100%" border="0" cellspacing="1" cellpadding="2">
                      <tr>  
                        <td width="26%" height="2">  
                          <div align="right"><b><font color="#CC0000"><b>*</b></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Employee  
                            No : </font></b></div>
                        </td>
                        <td width="74%" height="2"> <%	if Mid(Session("sAuthor"),19,3) = "AED" and sMode <> "EDIT" Then%>  
                          <input type="text" name="sEmpno" size="15" maxlength="12" value="<%=sEmpno%>" <%=sbread%>>
                          <%  
else
Response.write sEmpno
End if
%> <%Show_Err(sEmpno)%> </td>
                      </tr>
<%if Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Employee No Data  
                          (Character) : Not more than 12 Characters ,Example :  
                          00001. </font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Division  
                            : </font></b></div>
                        </td>
                        <td width="74%">
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                          <select name="sDivcode">
                            <%create_pDivision(sDivcode)%>  
                          </select>
                          <%=img_sDivision%> <FONT COLOR="#006699"><% Else %></FONT>
                          <%List_Division(sDivcode)%> <FONT COLOR="#006699"><%End if%></FONT>  
						</td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Division of Employee.  
                          Select from List.</font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Department  
                            : </font></b></div>
                        </td>
                        <td width="74%">  
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                          <select name="sDepcode">
                            <%create_pDepartment(sDepcode)%>  
                          </select>
                          <%=img_sDepartment%> <FONT COLOR="#006699"><% Else%></FONT>
                          <%List_Department(sDepcode)%> <FONT COLOR="#006699"><%End if%></FONT> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Department of Employee.  
                          Select from List.</font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Section  
                            : </font></b></div>
                        </td>
                        <td width="74%">  
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                          <select name="sSeccode">
                            <%create_pSection(sSeccode)%>  
                          </select>
                          <%=img_sSection%> <FONT COLOR="#006699"><% Else%></FONT>
                          <%List_Section(sSeccode)%> <FONT COLOR="#006699"><%End if%></FONT> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Section of Employee.  
                          Select from List.</font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">  
                            </font><font color="#CC0000"><b>*</b></font> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Position  
                            : </font></b></div>
                        </td>
                        <td width="74%">  
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                          <select name="sPoscode">
                            <%create_pPosition(sPoscode)%>  
                          </select>
                          <%=img_sPosition%> <FONT COLOR="#006699"><% Else%></FONT>
                          <%List_Position(sPoscode)%> <FONT COLOR="#006699"><%End if%></FONT> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Position of Employee.  
                          Select from List.</font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">  
                            </font><font color="#CC0000"><b>*</b></font> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Status  
                            : </font></b></div>
                        </td>
                        <td width="74%">
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                          <select name="sStatuscode">
                            <%create_pStatus(sStatuscode)%>  
                          </select>
                          <%=img_sStatus%> <font color="#006699"><% Else%></font>
                          <%List_Status(sStatuscode)%> <font color="#006699"><%End if%></font> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Status of Employee.  
                          Select from List.</font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"><b>*</b></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Name  
                            - Surname : <br>
                            </font></b></div>
                        </td>
                        <td width="74%"> <%if Mid(Session("sAuthor"),1,3) = "999" Or Mid(Session("sAuthor"),19,3) = "AED" Then%>  
                          <input type="text" name="sName" value="<%=sName%>" SIZE="40" maxlength="60">
                          <%Else%> <%=sName%> <%End if%> <%=img_sName%> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333"> Employee Name and  
                          Surname(English Language). Not more than 60 Characters  
                          . </font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Name  
                            - Surname : <br>
                            </font></b></div>
                        </td>
                        <td width="74%"> <%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>  
                          <input type="text" name="sTName"  value="<%=sTName%>">
                          <%Else%> <%=sTName%> <%End if%> <%=img_sTName%> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333"> Employee Name and  
                          Surname (Thai Language). Not more than 60 Characters  
                          </font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"><b>*</b></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Login  
                            Name : </font></b></div>
                        </td>
                        <td width="74%"> <%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>  
                          <input type="text" name="sLogin" size="12" maxlength="12" value="<%=sLogin%>">
                          <%Else%> <%=sLogin%> <%End if%> <%=img_sLogin%> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Name for Login (Character).  
                          Not more than 12 Characters .</font></td>
                      </tr>
<%End if%>

<%if Mid(Session("sAuthor"),1,3) <> "999" Then%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"><b>*</b></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Old Password  
                            : </font></b></div>
                        </td>
                        <td width="74%">  
                          <input type="password" name="sOldPassword" size="12" maxlength="12">
                          <%=img_sOldPassword%> </td>
                      </tr>

<%end if%>


                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"><b>*</b></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">New Password  
                            : </font></b></div>
                        </td>
                        <td width="74%">  
                          <input type="password" name="sPassword" size="12" maxlength="12"  
							<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
								 value="<%=sPassword%>"
							<%End if%> >
							<%if Mid(Session("sAuthor"),1,3) <> "999" then%>3-12 Characters .<%End if%>
                          <%=img_sPassword%> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Password Login (Character).  
                          Between 3-12 Characters .</font></td>
                      </tr>
<%End if%>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"><b>*</b></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">re-New Password  
                            : </font></b></div>
                        </td>
                        <td width="74%">  
                          <input type="password" name="sPassword1" size="12" maxlength="12"  
							<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
								 value="<%=sPassword%>"
							<%End if%>>
							<%if Mid(Session("sAuthor"),1,3) <> "999" then%>3-12 Characters .<%End if%>
                          <%=img_sPassword%> </td>
                      </tr>
<%if Mid(Session("sAuthor"),1,3) = "999"  Or Mid(Session("sAuthor"),19,3) = "AED" Then%>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">Re-Type your Password  
                          Login (Character). Between 3-12 Characters .</font></td>
                      </tr>
<%End if%>
                     <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">E-Mail
                             : </font></b></div>
                        </td>
                        <td width="74%">
                          <input type="text" name="sRecMail" size="30" maxlength="60" value="<%=sRecMail%>"><%=Img_RecMail%>
                          </td>
                      </tr>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC3333">
                         Example : admin@amec.ina.melco.co.jp </font></td>
                      </tr>
                      <tr>  
                        <td width="26%">  
                          <div align="right"><b><font color="#CC0000"></font>  
                            <font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp;
                            </font></b></div>
                        </td>
                        <td width="74%">  
                          <input type="checkbox" name="cRecMail"  value="Y" <%IF cRecMail ="Y" then Response.write "checked"%>> Receive Approve E-Mail
						  </td>
                      </tr>

                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"><font color="#CC0000"><b>*</b></font>  
                          Required Field.</td>
                      </tr>
                      <tr>  
                        <td width="26%">&nbsp;</td>
                        <td width="74%"> <% if sMode = "EDIT" Then %>  
                          <input type=HIDDEN name=sEmpno value='<%=sEmpno%>'>
                          <% end if %>  
                          <input type=HIDDEN name=EVENT value='<%=sMode%>'>
                          <input type=HIDDEN name=sDivcode value='<%=sDivcode%>'>
                          <input type=HIDDEN name=sDepcode value='<%=sDepcode%>'>
                          <input type=HIDDEN name=sSeccode value='<%=sSeccode%>'>
                          <input type=HIDDEN name=sPoscode value='<%=sPoscode%>'>
                          <input type=HIDDEN name=sStatuscode value='<%=sStatuscode%>'>
                          <input type=HIDDEN name=sName value='<%=sName%>'>
                          <input type=HIDDEN name=sTName value='<%=sTName%>'>
                          <input type=HIDDEN name=sLogin value='<%=sLogin%>'>
                          <input type=HIDDEN name=sOldPassword1 value='<%=sOldPassword1%>'>
                          <input type=HIDDEN name=clevel value='admin'>
                          <input type=HIDDEN name=chk value='<%=Request("chk")%>'>
                          <input type="submit" name="Submit" value="Submit (Save)">
                          <input type="reset" name="Reset" value="Reset">
                        </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <tr align="center">  
                <td colspan="2"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="3" COLOR="#666666"><B><FONT COLOR="#FFCC66">..</FONT></B></FONT>  
                </td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a>  
              | <a href="../security/security.asp">Security Manager</a></font></p>
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
            </b><font size="1">700/86 Amata Nakorn Industrial Estate,Moo 6,Bangna-Trad Rd.,<br>  
T.Don Hua Roh,Muang District,Chonburi 20000 <br>
Tel:(038) 213-170-80 Fax: (038) 213-184-6</font></font></td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
<br>
<% End Sub %>
<%			
 ' Response.write ("sEmpno = " & sEmpno)
 ' Response.write (" sDivcode = " & sDivcode)
 ' Response.write (" sDepcode = " & sDepcode)
 ' Response.write (" sSeccode = " & sSeccode)
 ' Response.write (" sPoscode = " & sPoscode)
 ' Response.write (" sName = " & sName)
 ' Response.write (" sSurname = " & sSurname)
 ' Response.write (" sTName = " & sTName)
 ' Response.write (" sTSurname = " & sTSurname)
 ' Response.write (" sLogin  = " & sLogin)
 ' Response.write ( "sPassword = " & sPassword)
 ' Response.write (" sPassword1 = " & sPassword1)  
%>

<%
Sub Update_Employee(sUpdateMode)

  sMode = sUpdateMode
  sEmpno = Ucase(Request("sEmpno"))
  sDivcode = Request("sDivcode")  
  sDepcode = Request("sDepcode")
  sSeccode = Request("sSeccode")
  sPoscode = Request("sPoscode")  
  sStatuscode = Request("sStatuscode")  
  sName = Ucase(Request("sName"))
  sSurname = Request("sSurname")
  sTName = Request("sTName")
  sTSurname = Request("sTSurname")
  sLogin = Ucase(Request("sLogin"))
  sPassword = Request("sPassword")
  sOldPassword2 = Request("sOldPassword")
  sOldPassword1 = Request("sOldPassword1")
  sPassword1 = Request("sPassword1")
  sUpdateMode =	Request("nMode")
  sRecMail=Request("sRecMail")
  IF Request("cRecMail")<>"" then
  	cRecMail=Request("cRecMail")
  Else
	  cRecMail="N"
  End IF
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	' Check if the End date is a valid date provided it is not empty

	If (Trim(sDivcode) <> "00" ) Then
		If (Trim(sDepcode) <> "00") Then
			If (Trim(sSeccode) <> "00") Then
				sSeccode = sSeccode  
			Else
				sSeccode = "00"					
			End If
			sDepcode = sDepcode
		Else  
			sDepcode = "00"
		End If
		sDivcode = sDivcode
	Else
		sDivcode="00"
	End If


	If (Trim(sDivcode) = "00" and Trim(sDepcode) <> "00" and Trim(sSeccode) = "00") Then
		bError = True		
		img_sDivision = sErrImage
    End If

	If (Trim(sDivcode) = "00" and Trim(sDepcode) <> "00" and Trim(sSeccode) <> "00") Then
		bError = True		
		img_sDivision = sErrImage
    End If
  
	If (Trim(sDivcode) = "00" and Trim(sDepcode) = "00" and Trim(sSeccode) <> "00") Then
		bError = True		
		img_sDivision = sErrImage
		img_sDepartment = sErrImage
    End If

	If (Trim(sDivcode) <> "00" and Trim(sDepcode) = "00" and Trim(sSeccode) <> "00") Then
		bError = True
		img_sDepartment = sErrImage
 	 End If



	If (Trim(sEmpno) = "") Then
		bError = True
		img_sEmpno = sErrImage
	End If


'	If (Trim(sName) = "") Then
'		bError = True
'		img_sName = sErrImage
'	End If

'	If (Trim(sStatuscode) = "") Then
'		bError = True
'		img_sSurname = sErrImage
'	End If

'	If (Trim(sTName) = "") Then
'		bError = True
'		img_sTName = sErrImage
'	End If

'	If (Trim(sTSurname) = "") Then
'		bError = True
'		img_sTSurname = sErrImage
'	End If
'IF sMode="ADD" then
	If (Trim(sPassword) = "") Then
		bError = True
		img_sPassword = sErrImage
	End If

	If (Trim(sPassword1) = "") Then
		bError = True
		img_sPassword = sErrImage
	End If

	if len(sPassword) < 3 Then
		bError = True
		img_sPassword = sErrImage

	End If

	if trim(sPassword) <> trim(sPassword1) then
		bError = True
		img_sPassword = sErrImage
	End If
	if trim(sOldPassword1) <> trim(sOldPassword2) then
		bError = True
		img_sOldPassword = sErrImage
	End If

'End if

	If Trim(sRecMail) = "" and Trim(cRecMail)<>"N"  Then
		bError = True
		img_RecMail = sErrImage
	End IF

	If bError = False Then
	  ' Create a recordset
	    	Set Rs = Server.CreateObject("ADODB.RecordSet")
		If sMode = "EDIT" Then		   
          sSQL = "SELECT * FROM  aEmployee WHERE sEmpno = '" & sEmpno & "'" 		   
          Rs.Open sSQL, conn, adOpenStatic, adLockPessimistic, adCmdText
		Else
         sSQL = "SELECT * FROM aEmployee"  
		  Rs.Open sSQL, conn, adOpenStatic, adLockPessimistic, adCmdText
		  Rs.AddNew
		End If
		
'		RS("sEmpno") = sEmpno
'		RS("sDivcode") = sDivcode
'		RS("sDepcode") = sDepcode
'		RS("sSeccode") = sSeccode
'		RS("sPoscode") = sPoscode
'		RS("cStatus") = Request("sStatuscode")
'		RS("sName") = sName
'		RS("sSurname") = sSurname
'		RS("sTName") = sTName
'		RS("sTSurname") = sTSurname
'		RS("sLogin") = sLogin
		RS("sPassword") = sPassword
		RS("cRecMail")=cRecMail
		RS("sRecMail")=sRecMail
		RS.Update
			
		RS.Close
		Set RS = Nothing
		
'		If sMode = "ADD" Then		   
'			Set_Authorize(sEmpno)
'		End If

'		sEmpno = ""
'		Response.Redirect  "pEmployee.asp"
'		Response.Redirect  "../security/security.asp"
'		Response.Redirect  "pEmployee.asp?keywords='"&sEmpno&"'"
		sTatus="del"
'		if Request("chk")="admin" then
			Response.Redirect  "../security/security.asp"
'		else
'		  Response.Redirect("pEmployee.asp?sTatus=" & status)
'		End if
	Else
		Show_Form()
	End If	' if Error
End Sub		
%>
 <%
Sub Show_Err(Mesg)
      if not IsEmpty(request("EVENT")) then     
            if Trim(Mesg) = "" then
%> <img src="../images/icerror.gif"> <%
	         end if      
      end if  
End Sub
%>

<%
Sub Set_Authorize(sEmpno)
		Dim sProgcode,sAuthor,sNote,sAuthor2,sAuthor3,sAuthor4,sAuthor5,sAuthor6,sAuthor7,sAuthor8,sAuthor9,sAuthor10,sAuthor11
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Rs.Open "aAuthorize", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		Rs.AddNew
		'Response.Write		 sSQL
		sProgcode = "000"
		sNote = ""
		sAuthor2 = "NOT"
		sAuthor3 = "NOT"
		sAuthor4 = "NOT"
		sAuthor5 = "NOT"
		sAuthor6 = "NOT"
		sAuthor7 = "NOT"
		sAuthor8 = "NOT"
		sAuthor9 = "NOT"
		sAuthor11 = "NOT"
		Rs("sEmpno") = sEmpno
		Rs("sProgcode") = sProgcode
		Rs("sAuthor") = "0" & show_pos(sempno) & sAuthor2 & sAuthor3 & sAuthor4 & sAuthor5 & sAuthor6 & sAuthor7 & sAuthor8 & sAuthor9 & Show_Div(sempno) & sAuthor11
		Rs("sNote") = sNote
	  Rs.Update		

	  Rs.Close
	  Set Rs = Nothing
End Sub
%>
