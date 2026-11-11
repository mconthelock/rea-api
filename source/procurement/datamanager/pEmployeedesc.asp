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
	Dim img_sDivision, img_sDepartment, img_sSection, img_sPosition
	Dim sDivision, sDepartment, sSection, sPosition
	Dim sDivcode, sDepcode, sSeccode, sPoscode
	Dim sEmpno, sDivc, sDepc, sSecc, sPO, sName, sSurname, sTName, sTSurname
	Dim sLogin, sPassword, sPassword1
	Dim bError, sbread
'	Dim sMode
	Dim sErrImage
	bError = False
	sErrImage = "<IMG SRC='../images/icerror.gif'>"
sScriptPath = Request.ServerVariables("SCRIPT_NAME")

sEmpno=Request("sEmpno")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM vaEmployee WHERE sEmpno ='" & sEmpno & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sEmpno = Rs("sEmpno")
		sDivcode = Rs("sDivcode")
		sDivision=RS("sDivision")
		sDepcode = Rs("sDepcode")
		sDepartment=RS("sDepartment")
		sSeccode = Rs("sSeccode")
		sSection=Rs("sSection")
		sPoscode = Rs("sPoscode")
		sPosition=RS("sPosition")
		sName = Rs("sName")
		sSurname = Rs("sSurname")
		sTName = Rs("sTName")
		sTSurname = Rs("sTSurname")
		sLogin = Rs("sLogin")
		sPassword = RS("sPassword")
		sbread="ReadOnly"
	Else
		sMode = "ADD"
	End If

%>	

<%
' This routines presents the form to the user. Depending on the mode, the
' form fields are filled with data or blank.
	' Set the title of the page depending on if an error occurred during
	' form submission or not
%>
<html>
<head>
<title><%=sMode%> Employee</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
</SCRIPT>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="Form_Load()" >
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
                  <table width="100%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td width="25%" height="2"> 
                        <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Employee 
                          ID : </font></b></div>
                      </td>
                      <td width="75%" height="2"> <!--	<input type="text" name="sEmpno" size="15" maxlength="12" value="<%=sEmpno%>" <%=sbread%>> --> 
                        <%=sEmpno%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="25%"> 
                        <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Division 
                          : </font></b></div>
                      </td>
                      <td width="75%"> <!--	<input type="text" name="sDivcode" size="80" maxlength="12" value="<%List_Division(sDivcode)%>" <%=sbread%>> --> 
                        <%List_Division(sDivcode)%> 
                        <font color="#006699"></font> </td>
                    </tr>
                    <tr> 
                      <td width="25%"> 
                        <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Department 
                          : </font></b></div>
                      </td>
                      <td width="75%"> <!--                          <INPUT TYPE="text" NAME="sDepcode" SIZE="80" MAXLENGTH="12" VALUE="<%List_Department(sDepcode)%>" <%=sbread%>> --> 
                        <%List_Department(sDepcode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="25%"> 
                        <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Section 
                          : </font></b></div>
                      </td>
                      <td width="75%"> <!--                          <INPUT TYPE="text" NAME="sSeccode" SIZE="80" MAXLENGTH="12" VALUE="<%List_Section(sSeccode)%>" <%=sbread%>> --> 
                        <%List_Section(sSeccode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="25%"> 
                        <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Position 
                          : </font></b></div>
                      </td>
                      <td width="75%"> <!--                          <INPUT TYPE="text" NAME="sPoscode" SIZE="80" MAXLENGTH="12" VALUE="<%List_Position(sPoscode)%>" <%=sbread%>> --> 
                        <%List_Position(sPoscode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="25%"> 
                        <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Name 
                          &lt; English &gt; : </font></b></div>
                      </td>
                      <td width="75%"> <!--                          <input type="text" name="sName" value="<%=sName%>" SIZE="40"> --> 
                        <%=sName%> 
                      </td>
                    </tr>
                    <!--              <tr> 
                        <td width="25%"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Surname 
                            &lt; English &gt; </font></div>
                        </td>
                        <td width="75%"> 
                          <input type="text" name="sSurname" value="<%=sSurname%>">
                           </td>
                      </tr> --> 
                    <tr> 
                      <td width="25%"> 
                        <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Name 
                          &lt; Thai &gt; : </font></b></div>
                      </td>
                      <td width="75%"> <!--                          <input type="text" name="sTName"  value="<%=sTName%>"> --> 
                        <%=sTName%> 
                      </td>
                    </tr>
                    <!--                      <tr> 
                        <td width="25%"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Surname 
                            &lt; Thai &gt;</font></div>
                        </td>
                        <td width="75%"> 
                          <input type="text" name="sTSurname" value="<%=sTSurname%>">
                          </td>
                      </tr> --> <!--                      <tr> 
                        <td width="25%"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Login 
                            Name</font></div>
                        </td>
                        <td width="75%"> 
                          <input type="text" name="sLogin" size="12" maxlength="12" value="<%=sLogin%>">
                          <%=img_sLogin%> </td>
                      </tr>
                      <tr> 
                        <td width="25%"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Password 
                            </font></div>
                        </td>
                        <td width="75%"> 
                          <input type="password" name="sPassword" size="12" maxlength="12" Value='<%=sPassword%>'>
                          <%=img_sPassword%> </td>
                      </tr>
                      <tr> 
                        <td width="25%"> 
                          <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">re-Password 
                            </font></div>
                        </td>
                        <td width="75%"> 
                          <input type="password" name="sPassword1" size="12" maxlength="12" Value='<%=sPassword%>'>
                          <%=img_sPassword%> </td>
                      </tr> --> 
                    <tr> 
                      <td width="25%">&nbsp;</td>
                      <td width="75%"> </td>
                    </tr>
                  </table>
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
            <!--#include file="../inc/copyRight.inc"-->
		   </td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b></b></font></td>
  </tr>
</table>
<br>
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


