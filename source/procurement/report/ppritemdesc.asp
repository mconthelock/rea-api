<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
'on error resume next
  	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprno,nitem,sprodcode,sprodname,svendcode,sproddesc,nquality,nprice,namount,scurcode,susefor,scomment,cstatus
	Dim img_nitem, img_Prno,sbread,nstep,Item,sexptype,img_nQuality,img_nPrice,img_nAmount,img_sVendcode,img_cStatus
	Dim sUnitcode,sQuotation,sImagespec,snote
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
  	sSQL = "SELECT * FROM ppritem WHERE nitem='" & nitem & "'" &_
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
		sNote=RS("sNote")
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
                    <table width="98%" border="0" cellspacing="1" cellpadding="1">
                      <tr> 
                        <td width="30%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                          No. : &nbsp;</font></b></td>
                        <td width="70%"> <%=sPrno%></td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Code : &nbsp;</font></b></td>
                        <td width="75%"> 
									<%=sProdcode%> </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Name(ENG) : &nbsp;</font></b></td>
                        <td width="75%"> 
                          <%=sProdname%>
                        </td>
                      </tr>
						<%Show_Prod()%>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                          : &nbsp;</font></b></td>
                        <td width="75%"> 
                          <%Show_Vender(sVendcode)%>
                         </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                          Description : &nbsp;</font></b></td>
                        <td width="75%"> 
                      <%=sProddesc%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quantity 
                          : &nbsp; </font></b></td>
                        <td width="75%"> 
                          <%=FormatNumber(nQuality)%>
                           <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"><b> &nbsp; </b></font>
                            <%Show_Unit(sUnitcode)%> 
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UPrice 
                          : &nbsp;</font></b></td>
                        <td width="75%"> 
                          <%=FormatNumber(nPrice, 3)%> &nbsp; <%Show_Currency(sCurcode)%> 
                       
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Amount : &nbsp;</font></td>
                        <td width="75%"> 
                          <%=FormatNumber(nAmount, 3)%>&nbsp; <%Show_Currency(sCurcode)%> 
                        </td>
                      </tr>  
<!--                       <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">UseFor Section
                          : &nbsp;</font></b></td>
                        <td width="75%" height="33"> 
          					<%=sUsefor%>
                        </td>
                      </tr> -->
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Comment
                          : &nbsp;</font></b></td>
                        <td width="75%" height="33"> 
          					<%=sComment%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="25%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Quotation No. 
                          : &nbsp;</font></b></td>
                        <td width="75%" height="33"> 
          					<%=sQuotation%>
                        </td>
                      </tr>
                       <tr> 
                        <td width="25%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Image : &nbsp;</font></b></td>
                        <td width="75%"><a href="../datamanager/ProductImages/<%=sImagespec %>" target="_Blank"><%=sImagespec %></a><br>

                        </td>
                      </tr>
                      <tr> 
                        <td width="19%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Note 
                          : &nbsp;</font></b></td>
                        <td width="81%"> 
                          <%=sNote%>
                        </td>
                      </tr> 
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
             <!--#include file="../inc/copyRight.inc"--></td>
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

			Rs2.MoveNext
		Loop
	End IF
	RS2.Close
	Set RS2=Nothing
End Sub

Sub Show_Prod()
Dim sSQL2,RS2
	sSQL2="Select * from pProduct where sprodcode='" & sprodcode & "'" &_
	" and sExptype='" & Request("sExptype") "'" 
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF Not Rs2.EOF then
%>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product Name(THA)
                        :  &nbsp;</font></b></td>
                      <td width="74%"> <% =RS2("sTprodname")%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Type 
                        :  &nbsp;</font></b></td>
                      <td width="74%"> <% Show_Pexptype(RS2("sexptype"))%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Group 
                        :  &nbsp;</font></b></td>
                      <td width="74%"> <% Show_Pcattype(RS2("sCattype"))%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Sub-Group 
                        :  &nbsp;</font></b></td>
                      <td width="74%"> <% Show_PCatalog(RS2("sCatID"))%> 
                      </td>
                    </tr>
<%		
	End if	
End Sub
%>
<%
Sub Show_pCatalog(sCatID)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCatalog Where sCatID='" & sCatID & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sCatID") = sCatID Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write(Rs2("sCatalog"))
			Rs2.MoveNext
		Loop
	End IF
End Sub
Sub Show_pCattype(sCattype)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCattype Where sCattype='" & sCattype & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sCattype") = sCattype Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write(Rs2("sCatdesc"))
			Rs2.MoveNext
		Loop
	End IF
End Sub
%>
