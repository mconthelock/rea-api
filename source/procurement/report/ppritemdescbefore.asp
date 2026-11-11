<% Option Explicit %>
<%
           ' Author : Ekkapan Kongchoo
%>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
	' Author: Phornthep Khongsathian

%>	

<%
on error resume next
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprno,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Prno,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode,img_cStatus
	Dim sUnitcode,sQuotation,sImagespec
		sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		nitem=Request("nitem")
		sprno=Request("sprno")
		sExptype = Request("sExptype")
		sVendcode = Request("sVendcode")
		sProdcode = Request("sProdcode")



	If IsEmpty(Request.Form("EVENT")) Then
		If  nitem <> "" Then
			sMode = "EDIT"
			if Request("status")<>"change" then
				Get_Pritem(sPrno)
			end if
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Pritem(Request.Form("EVENT"))
	End If		
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Pritem(sPrno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  	sSQL = "SELECT * FROM ppritem2 WHERE nitem='" & nitem & "'" &_
	"and sprno='" & sprno & "'"
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sPrno=RS("sPrno")
		nItem = Rs("nItem")
		sProdcode=RS("sProdcode")
		sProdname=RS("sProdname")
		sVendcode= Rs("Svendcode")
		sProddesc=RS("sProddesc")
		nQuality=RS("nQuality")
		nPrice=RS("nPrice")
		nAmount=RS("nAmount")
		sCurcode=RS("sCurcode")
		sUseFor=RS("sUseFor")
		sQuotation=RS("sQuatation")
		sComment=RS("sComment")
		sUnitcode=RS("sUnitcode")
		cStatus=RS("cStatus")
		simagespec=Rs("sImagespec")
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
<title><%=sMode%> P/R Item</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
<Script Langauge="JavaScript">
<!---
function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
            WORK FLOW MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="25%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> 
                  P/R Item</font></b></font></td>
                <td width="75%" bgcolor="#DB4B37" align="right"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                          No. : &nbsp;</font></b></td>
                        <td width="81%"> <%=sPrno%></td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Item 
                          No. : &nbsp;</font></b></td>
                        <td width="81%"><%=nItem%></td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Code : &nbsp;</font></b></td>
                        <td width="81%"> 
									<%=sProdcode%> </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Name : &nbsp;</font></b></td>
                        <td width="81%"> 
                          <%=sProdname%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                          : &nbsp;</font></b></td>
                        <td width="81%"> 
                          <%Show_Vender(sVendcode)%>
                         </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Description : &nbsp;</font></b></td>
                        <td width="81%"> 
                      <%=sProddesc%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quantity 
                          :&nbsp; </font></b></td>
                        <td width="81%"> 
                          <%=FormatNumber(nQuality)%>
                           <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b> &nbsp; </b></font>
                            <%Show_Unit(sUnitcode)%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UPrice 
                          : &nbsp;</font></b></td>
                        <td width="81%"> 
                          <%=FormatNumber(nPrice, 3)%> &nbsp; <%Show_Currency(sCurcode)%> 
                       
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Amount : &nbsp;</font></td>
                        <td width="81%"> 
                          <%=FormatNumber(nAmount, 3)%>&nbsp; <%Show_Currency(sCurcode)%> 
                        </td>
                      </tr>  
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UseFor Sec.
                          : &nbsp;</font></b></td>
                        <td width="81%" height="33"> 
          					<%=sUsefor%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quotation No. 
                          : </font></b></td>
                        <td width="81%" height="33"> 
          					<%=sQuotation%>
                        </td>
                      </tr>
                       <tr> 
                        <td width="19%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Image :</font></b></td>
                        <td width="81%"><a href="../datamanager/ProductImages/<%=sImagespec %>" target="_Blank"><%=sImagespec %></a><br>

                        </td>
                      </tr>
<!--                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Comment 
                          : &nbsp;</font></b></td>
                        <td width="81%"> 
                          <%=sComment%>
                        </td>
                      </tr> -->
                    </table>
            
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../security/images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../security/images/y04.jpg" width="12" height="12"></td>
        </tr> 
      </table><a href="Javascript:window.close()"> <center>Close</center> </a>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr align="center"> 
          <td> <br>
            <hr noshade size="1">
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#666666"><b><font size="1">MITSUBISHI 
            ELEVATOR ASIA CO.,LTD.</font><br>
            </b><font size="1">700/86 Amatacity Chonburi Industrial Estate, Moo 6,Debaratana Road,<br>
Donhuaroh Sub-district, Mueang Chonburi District, Chonburi Province 20000<br> 
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

<% End Sub %>
<%
Sub show_file()
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from ppritemdesc Where sprno='" & sprno & "'" &_
		" and nitem='" & nitem & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
%><a href="../workflow/ppritemdescimages/<%=RS2("sfile")%>" target="_Blank">
<%					Response.Write(Rs2("sfile") & "</a>&nbsp;&nbsp;&nbsp;")
					Response.write(RS2("sdesc") & "&nbsp;&nbsp;&nbsp;&nbsp;")
%>
<%
			Rs2.MoveNext
		Loop
	End IF
End Sub
%>

