<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	'########## DERMPAN ##########
	'I change FormatNumber(x, 2) to FormatNumber(x, 3).
	'And increase nPrice column in pproduction to 16,3
	'########## END DERMPAN ##########

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sedesc,nprice,simage,dadddate,sempno
	Dim img_sprodcode	,img_sexptype,img_svendcode,img_nprice,sbread,sAcccode,sunitcode,scurcode
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year ,sErrImage,sCatID
	Dim sCattype,hazard,hazardstatus
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sprodcode=Request.QueryString("sprodcode")
	sexptype=Request.QueryString("sexptype")
	svendcode=Request.QueryString("svendcode")
	sCatID=Request.QueryString("sCatID")
	sCattype=Request("sCattype")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
if sExptype<> "" and sCattype<>"" and sCatID<>"" then
	  sSQL = "SELECT PPRODUCT.*,HAZARDTYPEMST.HAZARDTYPE FROM PPRODUCT LEFT JOIN HAZARDTYPEMST ON PPRODUCT.hazardno = HAZARDTYPEMST.hazardno WHERE sprodcode='" & sprodcode & "'" &_
	" and sexptype='" & sexptype & "'" &_
	" and sCattype='" & sCattype & "'" &_
	" and sCatID='" & sCatID & "'"
elseif sExptype<> "" and sCattype<> "" and sCatID="" then
	  sSQL = "SELECT PPRODUCT.*,HAZARDTYPEMST.HAZARDTYPE FROM PPRODUCT LEFT JOIN HAZARDTYPEMST ON PPRODUCT.hazardno = HAZARDTYPEMST.hazardno WHERE sprodcode='" & sprodcode & "'" &_
	" and (sexptype='" & sexptype & "'" &_
	" and sCattype='" & sCattype & "')"
else
	  sSQL = "SELECT PPRODUCT.*,HAZARDTYPEMST.HAZARDTYPE FROM PPRODUCT LEFT JOIN HAZARDTYPEMST ON PPRODUCT.hazardno = HAZARDTYPEMST.hazardno WHERE sprodcode='" & sprodcode & "'" &_
	" and sExptype='" & sExptype & "'"
End if

   Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1

  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sprodcode = Rs("sprodcode")
		sexptype = Rs("sexptype")
		sCattype=Rs("sCattype")
		sCatID=Rs("sCatID")
		svendcode = Rs("svendcode")
		sprodid = Rs("sprodid")
		stprodname = Rs("stprodname")
		seprodname = Rs("seprodname")
		smodel = Rs("smodel")
		sAcccode = Rs("sAcccode")
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
		sempno = RS("sempno")
		sbread = " readonly"
		hazard = rs("hazardtype")
		hazardstatus = rs("hazardstatus")
	End If
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
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Product </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2">
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product 
                        Code : </font></b></td>
                      <td width="44%"> <%=SPRODCODE%> 
					  </td>
					  <td width="30%" rowspan="18" valign="top" align="right">
					   
					  	<%
						'response.write "ttt = "&UCase(right(simage,3))
						If simage<> "" then 
						   if LCase(right(simage,3)) <> "pdf" then
						
						%>
							<!--a href="../datamanager/productimages/<%=d_year%>/<%=simage%>" target="_Blank"><%=simage%></a--> 
							<img src="../datamanager/ProductImages/<%=d_year%>/<%=simage%>" />
							
						<%
							end if
						end if
						
						%>
					  </td>
					  
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Type 
                        : </font></b></td>
                      <td  width="44%"> <% Show_Pexptype(sexptype)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Group 
                        : </font></b></td>
                      <td  width="44%"> <% Show_Pcattype(sCattype)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense Sub-Group 
                        : </font></b></td>
                      <td  width="44%"> <% Show_PCatalog(sCatID)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                        Name : </font></b></td>
                      <td  width="44%"> <%Show_Vender(svendcode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor's ProductID : </font></b></td>
                      <td  width="44%"> <%=sprodid%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Thai Product 
                        Name  : </font></b></td>
                      <td  width="44%"> <%=stprodname%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">English Product 
                        Name <english> : </font></b></td>
                      <td  width="44%"> <%=seprodname%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Model. 
                        : </font></b></td>
                      <td  width="44%"> <%=smodel%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Account Code. 
                        : </font></b></td>
                      <td  width="44%"> <%=sAcccode%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvet ica, sans-serif" size="2"  color="#CC3333">Thai 
                        Description : </font></b></td>
                      <td  width="44%"> <%=stdesc%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">English 
                        Description : </font></b></td>
                      <td  width="44%"> <%=sedesc%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">UPrice 
                        : </font></b></td>
                      <td  width="44%"> <%=FormatNumber(nprice,4)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">Unit 
                        : </font></b></td>
                      <td  width="44%"> <%show_unit(sunitcode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="30%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">Currency 
                        : </font></b></td>
                      <td  width="44%"> <%Show_currency(sCurcode)%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Date 
                        : </font></b></td>
                      <td  width="44%">&nbsp; <%   =d_day  %> 
                        &nbsp; <% response.write(getMonthName(d_month) & "&nbsp;")%> 
                        <font face="MS Sans Serif" size="2"></font> <%=d_year%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">By 
                        Emp No. : </font></b></td>
                      <td  width="44%"> <%=sempno%> 
                      </td>
                    </tr>
                    <tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Product Image
                         : </font></b></td>
                      <td  width="44%">
					
                      </td>
                    </tr>
					<tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvet ica, sans-serif" size="2"  color="#CC3333">HAZARDOUS STATUS  : </font></b></td>
                      <td  width="44%"> 
						<% if hazardstatus="1" then %> 
						Yes
						<%elseif hazardstatus="0" then %>
						No
						<%else%>
						<%end if%>
                      </td>
                    </tr>
					<tr> 
                      <td width="26%" align="Right"><b><font face="Verdana, Arial, Helvet ica, sans-serif" size="2"  color="#CC3333">HAZARDOUS SUBSTANCE  : </font></b></td>
                      <td  width="44%"> <%=hazard%> 
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;</td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../datamanager/Products.asp?status=del">Back To Product List</a> 
              | <a href="../datamanager/datamanager.asp">Data Manager</a> 
		<%	if Mid(Session("sAuthor"),10,3) = "AED" Then %>
			 | <a href="../datamanager/pproductedt.asp?sExptype_hid=0">Add 
              Next Product</a>
		<% End if%>
			</font></p>
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
