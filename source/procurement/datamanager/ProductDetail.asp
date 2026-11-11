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
	Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	Dim img_sprodcode	,img_sexptype,img_svendcode,img_nprice,sbread,sunitcode,scurcode
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
  sSQL = "SELECT * FROM PProduct WHERE sprodcode='" & sprodcode & "'" &_
"and sexptype='" & sexptype & "'"
   Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sprodcode = Rs("sprodcode")
		sexptype = Rs("sexptype")
		svendcode = Rs("svendcode")
		sprodid = Rs("sprodid")
		stprodname = Rs("stprodname")
		seprodname = Rs("seprodname")
		smodel = Rs("smodel")
		stdesc = Rs("stdesc")
		sedesc = Rs("sedesc")
		nprice = Rs("nprice")
		sUnitcode = Rs("sUnitcode")
		sCurcode = Rs("sCurcode")
		simage = Rs("simage")
		dadddate = Rs("dAdddate")
		d_day=day(dadddate)
		d_month=month(dadddate)
		d_year=year(dadddate)
		dDate=d_month & "/" & d_day & "/" & d_year
		sempno = RS("sempno")
		sbread = " readonly"
	End If
%>

 <html> <head> 
<title>Product Detail</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
</head>
<body bgcolor="#FFFFFF" leftmargin="3" topmargin="3" marginwidth="3" marginheight="3">
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
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Product </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <table width="100%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                        Code : </font></b></td>
                      <td width="70%"> <%=SPRODCODE%> 
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">ExpendType 
                        : </font></b></td>
                      <td width="70%"> <% Show_Pexptype(sexptype)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                        Name : </font></b></td>
                      <td width="70%"> <%Show_Vender(svendcode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor's Product 
                        ID : </font></b></td>
                      <td width="70%"> <%=sprodid%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                        Name <thai> : </font></b></td>
                      <td width="70%"> <%=stprodname%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                        Name <english> : </font></b></td>
                      <td width="70%"> <%=seprodname%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Model. 
                        : </font></b></td>
                      <td width="70%"> <%=smodel%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"  color="#CC3333">Thai 
                        Description : </font></b></td>
                      <td width="70%"> <%=stdesc%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">English 
                        Description : </font></b></td>
                      <td width="70%"> <%=sedesc%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">UPrice 
                        : </font></b></td>
                      <td width="70%"> <%=nprice%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">Unit 
                        : </font></b></td>
                      <td width="70%"> <%show_unit(sunitcode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">Currency 
                        : </font></b></td>
                      <td width="70%"> <%Show_currency(sCurcode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Image 
                        : </font></b></td>
                      <td width="70%"> <a href="../datamanager/productImages/<%=simage%>" target="_Blank"><%=simage%> </a>
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">AddDate 
                        : </font></b></td>
                      <td width="70%"><%'=Show_date(dDate)%> <%=day(dDate) & "/" & month(dDate) & "/" & year(dDate) %>
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Employee 
                        No. : </font></b></td>
                      <td width="70%"> <%=sempno%> 
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
            <DIV Align="Center"><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2"><A HREF="javascript:history.back()">Back</A> 
              | <A HREF="Javascript:window.close()"> Close </A></FONT></DIV>
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
