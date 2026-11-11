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
	Dim sexptype,sCatalog,sCattype,sCatID
	Dim img_sexptype, img_CatID, img_Cattype
	Dim cStatus, budget

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	sCatID=Request("sCatID")
	sCattype=Request("sCAttype")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")

	If IsEmpty(Request.Form("EVENT")) Then
		If  sCatID <> "" Then
			sMode = "EDIT"
			Get_Catalog(sCatID)
			Get_Budget(sCatID)
		Else
			sMode = "ADD"
			cStatus = 1
		End If	
		Show_Form()
	Else
		Update_Budget(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Catalog(sCatID)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT PCatalog.*, nvl2(epbudget.nbudget, epbudget.nbudget,0) nbudget FROM PCatalog,epbudget WHERE PCatalog.sCatID='" & sCatID & "' and PCatalog.sCatid = epbudget.sCatid(+) and cyear= '" & year(date)  & "'"
   sSQL = "SELECT * FROM PCatalog WHERE sCatID='" & sCatID & "'"
'	response.write sSQL
'	response.end
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sCatID = Rs("sCatID")
		sCattype = Rs("sCattype")
		sCatalog = Rs("sCatalog")
		cStatus = Rs("cStatus")
'		budget =  Rs("nbudget")
	Else
		sMode = "ADD"
	End If
	Rs.Close
End Sub	

Sub Get_Budget(sCatID)
	Dim sSQL, Rs

    Set Rs = Server.CreateObject("ADODB.RecordSet")
	sSQL = "SELECT nvl2(epbudget.nbudget, epbudget.nbudget,0) nbudget FROM epbudget WHERE sCatID='" & sCatID & "' and cyear= '" & year(date)  & "'"
	'response.write sSQL
    Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	if not Rs.EoF then
		budget = Rs(0)
	else	
		budget = 0
	end if
	Rs.Close
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
<title>Edit CATALOG </title>
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
          <td width="96%" align="left" valign="top" bgcolor="#FFD553"><b><img src="../images/plus.gif" width="9" height="9"> 
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="95%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Sub-Group in Expense Type  : Sub-Group in Expense Type Description of Products</font></b></font></td>
                <td width="5%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="30%" align="Right"><b>
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Sub-Group in Expense Type 
                          ID : </font></b></td>
                        <td width="70%"><%=sCatID%>
						<%
							'If sMode="EDIT" then
							'	Response.write sCatID
							'Else
						%> 
                          <!--<input type="text" name="sCatID" size="8" maxlength="6" value="<%'=sCatID%>">-->
                          <%'=img_CatID%> 
                          <%'End if%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="30%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Catalog 
                          Type : </font></b></td>
                        <td width="70%"> <%=sCattype%><%=sCatalog%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="30%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Year : </font></b></td>
                        <td width="70%">
						<select name="selY">
							<option value="<%=year(date)-1%>"><%=year(date)-1%></option>
							<option value="<%=year(date)%>" selected><%=year(date)%></option>
							<option value="<%=year(date)+1%>"><%=year(date)+1%></option>
						</select>
                        </td>
                      </tr>					  
                      <tr> 
                        <td width="30%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Budget : </font></b></td>
                        <td width="70%">
                          <input type="text" name="txtBudget" maxlength="15" size="30" value="<%=budget%>">
                        </td>
                      </tr>					  					  
                    </table>
                    <br>
                    <input type=hidden name=EVENT value="<%=sMode%>">
                    <input type=hidden name=status value="<%=Request("status")%>">
                    <%If sMode<> "ADD" and bError=false  then%> 
                    <input type=hidden name=sCatID value="<%=sCatID%>">
					<input type=hidden name=sCattype value="<%=sCattype%>">
					<input type=hidden name=sCatalog value="<%=sCatalog%>">
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
<br>
<br>
</body>
</html>

<% End Sub %>

<%
Sub Update_Budget(sUpdateMode)

	Dim sErrImage, Rs, sSQL
	Dim y, budget
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sCatID = Request.Form("sCatID")
		sCattype = Request.Form("sCattype")
		sCatalog = Request.Form("sCatalog")
		y = Request.Form("selY")
		budget = Request.Form("txtBudget")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty
	If (Trim(sCatID) = "") Then
		bError = True
		img_CatID = sErrImage
	End If
	If (Trim(sCattype) = "") Then
		bError = True
		img_Cattype = sErrImage
	End If
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		sSQL = "SELECT * FROM  epBudget WHERE sCatID='" & sCatID &"' and cyear= '" & y & "'"
		Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			
		If Rs.EoF then
			sSQL = "insert into epBudget(cyear, scatid, nbudget) values('" & y & "','" & sCatID &"' ,'" & budget & "')"
		Else
			sSQL = " update  epBudget set nbudget = '" & budget & "' WHERE sCatID='" & sCatID &"' and cyear= '" & y & "'"
		end If
'		response.write sSQL
'		response.end
	  	Rs.Close
		Rs.Open sSQL, conn
		
		Set Rs = Nothing
	     Response.Redirect "pCatalog.asp?sCattype=" & sCattype
	else
		Show_Form()
	End If	' if Error 
End Sub	
%>