<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	Dim img_sprodcode	,img_sexptype,img_svendcode,img_nprice,sbread,chkGroup
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year ,sErrImage
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sprodcode=Request.QueryString("sprodcode")
	sexptype=Request.QueryString("sexptype")
	svendcode=Request.QueryString("svendcode")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM PExpenstype Order By sExptype"

   Set Rs = Server.CreateObject("ADODB.RecordSet")
   Set Rs2 = Server.CreateObject("ADODB.RecordSet")
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText

%>

 <html> <head> 
<title>Show Product :: Computer</title>
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
<%		if Not Rs.EOF then
					Do While Not RS.EOF
					sExptype=RS("sExptype")
					chk_group()
 %>
  	                  <tr> 
  	                    <td colspan="3"  bgcolor="#FF9999">
					<%if chkgroup=0 then%>
						<a href="vpproductexp.asp?ListProduct=<%=sExptype%>">
					<%End if%>
						<b><%=RS("sExpdesc")%></b></a>
						</td>
  	                  </tr>
<% 			sSQL="Select * from pCattype where sExptype='" & RS("sExptype") & "'"
					 Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
						IF Not RS2.EOF then
							Do While Not Rs2.EOF
%>
		  	                  <tr> 
  		                    <td width="31%">
									<%If Not RS2.EOF then%>
											<a href='../datamanager/vpproduct2.asp?sCattype=<%=Server.UrlEncode(RS2("sCattype"))%>&sCatdesc=<%=Server.UrlEncode(RS2("sCatDesc"))%>'><%=RS2("sCatDesc")%></a><%RS2.MoveNext%>
									<%End If%>
								</td>
  		                    <td width="35%">
									<%If Not RS2.EOF then%>
											<a href='../datamanager/vpproduct2.asp?sCattype=<%=Server.UrlEncode(RS2("sCattype"))%>&sCatdesc=<%=Server.UrlEncode(RS2("sCatDesc"))%>'><%=RS2("sCatDesc")%></a><%RS2.MoveNext%>
									<%End If%>
								</td>
  		                    <td width="34%">
									<%If Not RS2.EOF then%>
											<a href='../datamanager/vpproduct2.asp?sCattype=<%=Server.UrlEncode(RS2("sCattype"))%>&sCatdesc=<%=Server.UrlEncode(RS2("sCatDesc"))%>'><%=RS2("sCatDesc")%></a><%RS2.MoveNext%>
									<%End If%>
								</td>
  		                  </tr>
<%
							 Loop
					
						End if
						RS2.Close
					RS.MoveNext
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
              | <a href="../datamanager/datamanager.asp">Data Manager</a> </font></p>
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
<%
Sub chk_Group()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCattype Where sExptype='" & sExptype & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
		chkgroup=1
	else
		chkGroup=0
	End IF
End Sub
%>
