<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim svendcode,stname,sename,staddress,seaddress,stelno,sfax,scontact,sposition,stnote,dadddate,sempno
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
		staddress = Rs("sTaddress")
		seaddress = Rs("sEaddress")
		stelno = Rs("sTelno")
		sfax = Rs("sFax")
		scontact = Rs("sContact")
		sposition = Rs("sPosition")
		stnote = Rs("sTnote")
		dadddate = Rs("dAdddate")
		d_day=day(dadddate)
		d_month=month(dadddate)
		d_year=year(dadddate)+543
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
          <td width="96%" align="center" valign="top"> 
            <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
              <TR> 
                <TD WIDTH="92%" BGCOLOR="#DB4B37"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="3" COLOR="#666666"><B><FONT COLOR="#FFFFFF" SIZE="2"> 
                  Vendor Description</FONT></B></FONT></TD>
                <TD WIDTH="8%" BGCOLOR="#8B2718" ALIGN="center"><IMG SRC="../images/iconhelp2.gif" WIDTH="26" HEIGHT="26"></TD>
              </TR>
              <TR BGCOLOR="#F8D7D1" ALIGN="center" VALIGN="top"> 
                <TD COLSPAN="2"> 
                  <FORM METHOD="post" ACTION="<%=sScriptPath%>">
                    <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
                      <TR BGCOLOR="#F8D7D1"> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Vendor 
                          Code : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"> <%=svendcode%></TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Thai 
                          Name : </FONT><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"> 
                          </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=stname%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Thai 
                          Short Name : </FONT><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"> 
                          </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=sstname%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">English 
                          Name : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=sename%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">English 
                          Short Name : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=ssename%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Thai 
                          Address : </FONT><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"> 
                          </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=staddress%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">English 
                          Address : </FONT><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"> 
                          </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=seaddress%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Tel. 
                          : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=stelno%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Fax. 
                          : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=sfax%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Email 
                          : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><A HREF="mailto:<%=sEmail%>"> 
                          <%=sEmail%></A></TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT SIZE="2" FACE="Verdana, Arial, Helvetica, sans-serif" COLOR="#CC3333">Contact 
                          : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=scontact%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT SIZE="2" FACE="Verdana, Arial, Helvetica, sans-serif" COLOR="#CC3333">Position 
                          : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=sposition%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Note 
                          : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=stnote%> </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT SIZE="2" FACE="Verdana, Arial, Helvetica, sans-serif" COLOR="#CC3333">AddDate 
                          : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><FONT FACE="MS Sans Serif" SIZE="2">วันที่&nbsp;</FONT> 
                          <%   =d_day  %> &nbsp; <% response.write(MonthName(d_month))%> 
                          <FONT FACE="MS Sans Serif" SIZE="2">พศ.</FONT> <%=d_year%> 
                        </TD>
                      </TR>
                      <TR> 
                        <TD WIDTH="35%" ALIGN="Right" VALIGN="TOP"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#CC3333">Employee 
                          No. : </FONT></TD>
                        <TD WIDTH="65%" VALIGN="TOP"><%=sempno%> </TD>
                      </TR>
                    </TABLE>
                  </FORM>
                </TD>
              </TR>
              <TR ALIGN="center"> 
                <TD COLSPAN="2">&nbsp; </TD>
              </TR>
            </TABLE>
            <DIV Align="Center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              <FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"><A HREF="javascript:history.back()">back</A> 
              | <A HREF="Javascript:window.close()"> Close </A></FONT></font></DIV>
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

