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
	Dim sCattype,sCatdesc,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	Dim img_sprodcode	,img_sexptype,img_svendcode,img_nprice,sbread,sExptype
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year ,sErrImage
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sCattype=Request.QueryString("sCattype")
	sExptype=Request("sExptype")
	'sCatdesc=Request.QueryString("sCatdesc")


'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
	sSQL="Select * from pCattype where sCattype='" & sCattype & "'"
   Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	sCatdesc=RS("sCatdesc")
  RS.Close

  sSQL = "SELECT * FROM PCatalog  WHERE sCattype='" & sCattype  & "'"



  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText

%>

 <html> <head> 
<title>Show Product :: Computer</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function CloseWindow_onClick() { //v2.0
  window.close();
}
//-->
</Script>
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
            DATA MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" height="181">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="181"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Product </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2" height="76"> 
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td colspan="3" bgcolor="#FF9999"><%=sCatdesc%></td>
                    </tr>
<%
				If Not RS.EOF then
					Do While Not Rs.EOF
%>
                    <tr> 
		                 
                      <td width="75%" > <%If Not RS.EOF then%> 
								<%=RS("sCatalog")%> 
                        <%End if%> </td>
 	                    <td width="25%">
						<a href="pproductedt.asp?sExptype=<%=sExptype%>&sCattype=<%=sCattype%>&sCatID=<%=RS("sCatID")%>&status=select&svendcode=<%=Request("sVendcode")%>" target="num">Select</a><%RS.MoveNext%>
						</td>
                    </tr>
<%
					Loop
				End if
%>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;</td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../datamanager/datamanager.asp">Data Manager</a>
			| <A HREF="Javascript:window.close()"> Close </A></font></p>
          </td>
          <td width="2%" height="181">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom" height="2"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom" height="2"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom" height="2"><img src="../security/images/y04.jpg" width="12" height="12"></td>
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
<br>
</body>
</html>
