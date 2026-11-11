<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim svendcode,stname,sename,saddress1,saddress2,stelno,sfax,scontact,sposition,stnote,dadddate,sempno
	Dim sSTname,sSEname,sEmail
	Dim img_sVendcode	,img_stname,img_sename,img_staddress,img_seaddress,img_stelno
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
svendcode=Request("svendcode")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	If IsEmpty(Request.Form("EVENT")) Then
		If  svendcode <> "" Then
			sMode = "EDIT"
				Get_Vender(svendcode)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Vender(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Vender(svendcode)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM Pvender WHERE svendcode='" & svendcode & "'"
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		svendcode = Rs("sVendcode")
		stname = Rs("sTname")
		sename = Rs("sEname")
		saddress1 = Rs("seaddress")
		saddress2 = Rs("staddress")
		stelno = Rs("sTelno")
		sfax = Rs("sFax")
		scontact = Rs("sContact")
		sposition = Rs("sPosition")
		stnote = Rs("sTnote")
		dadddate = Rs("dAdddate")
		d_day=day(dadddate)
		d_month=month(dadddate)
		d_year=year(dadddate)
		sempno = RS("sempno")
		sSTname=Rs("sSTname")
		sSEname=Rs("sSEname")
		sEmail=Rs("sEmail")
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
<title>Vender Description</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> 
                  Vendor Description</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="100%" border="0" cellspacing="1" cellpadding="2">
                      <tr bgcolor="#F8D7D1"> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                          Code : </font></b></td>
                        <td width="68%" VALIGN="TOP"> <%=svendcode%></td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Thai 
                          Name : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=stname%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Thai 
                          Short Name : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=sstname%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">English 
                          Name : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=sename%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">English 
                          Short Name : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=ssename%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Address 2 : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=saddress1%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Address 2 : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=saddress2%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Tel. 
                          : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=stelno%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Fax. 
                          : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=sfax%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Email 
                          : </font></b></td>
                        <td width="68%" VALIGN="TOP"><a href="mailto:<%=sEmail%>"> 
                          <%=sEmail%></a></td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Contact 
                          : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=scontact%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Position 
                          : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=sposition%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Note 
                          : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=stnote%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Date 
                          : </font></b></td>
                        <td width="68%" VALIGN="TOP"><font face="MS Sans Serif" size="2">&nbsp;</font> 
                          <%   =d_day  %> &nbsp; <% response.write(GetMonthName(d_month))%> 
                          <font face="MS Sans Serif" size="2">&nbsp;</font> <%=d_year%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="32%" align="Right" VALIGN="TOP"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">By 
                          Emp. No. : </font></b></td>
                        <td width="68%" VALIGN="TOP"><%=sempno%> 
                        </td>
                      </tr>
                    </table>
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

