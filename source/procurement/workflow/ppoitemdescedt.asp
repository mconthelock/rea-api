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

	Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	Dim sprodcode,sexptype,svendcode,sprodid,stprodname,seprodname,smodel,stdesc,sdesc,nprice,sfile,dadddate,sempno
	Dim img_sprodcode	,img_sexptype,img_svendcode,img_nprice,sbread,sCatID,Img_file
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year ,sErrImage
	Dim sEName,Mode,sCurcode,sUnitcode,sfilename,sAcccode,Max,spono
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
	sProdcode=Request.QueryString("sProdcode")
	sExptype=Request.QueryString("sExptype")
	sVendcode=Request.QueryString("sVendcode")
	Mode = Request.QueryString("Mode")

	If IsEmpty(Request.Form("EVENT")) Then
		If  sProdcode <> "" and sExptype <> "" Then
			sMode = "EDIT"
			Get_Product(sprodcode)
		Else
			sMode = "ADD"
		End If	
		Show_Form()
	Else
		Update_Product(Request.Form("EVENT"))
	End If	


' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Product(sprodcode)
	Dim sSQL, Rs
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
		sVendcode = Rs("svendcode")
		sprodid = Rs("sprodid")
		stprodname = Rs("stprodname")
		seprodname = Rs("seprodname")
		smodel = Rs("smodel")
		sAcccode=RS("sAcccode")
		stdesc = Rs("stdesc")
		sedesc = Rs("sedesc")
		sCatID = Rs("sCatID")
		nprice = Rs("nprice")
		simage = Rs("simage")
		dadddate = Rs("dAdddate")
		d_day=day(dadddate)
		d_month=month(dadddate)
		d_year=year(dadddate)
		sempno = RS("sempno")
		sCurcode=Rs("sCurcode")
		sUnitcode=Rs("sUnitcode")
		sbread = " readonly"
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
%> <html> <head> 
<title>Product</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link rel="stylesheet" href="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--


function OpenNewWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
-->
</SCRIPT>
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
            WORKFLOW MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="center" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"> 
                  P/R Item Description </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="../workflow/ppoitemdescSave.asp" name="frmName" enctype="multipart/form-data" >
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">

                      <tr> 
                        <td width="29%" align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"  color="#CC3333">Quotation 
                          : </font></b></td>
                        <td width="71%">
						<a href="../workflow/ppritemimages/<%=sfile%>"><%=sfile%></a>
                          <input type="file" name="sfile" size="18" maxlength="60" value="<%=sfile%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">Quotation (Quotation'name use a-z,A-Z,0-9,_) 
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Description: </font></b></td>
                        <td width="71%"> 
                          <input type="text" name="SDESC" size="35" maxlength="30" value="<%=sdesc%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="29%" align="Right">&nbsp;</td>
                        <td width="71%"><font color="#CC3333">
                          (Character) :Not more than 30 Characters 
                          .</font></td>
                      </tr>
                    </table>
                    <br>

					<input type="hidden"  name="sVendcode" value="<%=Request("sVendcode")%>">
                    <input type="hidden"  name="EVENT" value="<%=sMode%>">
					<input type="hidden"  name="status" value="<%=Request("status")%>">
<%if sMode="EDIT" and bError=false then%>
					<input type="hidden"  name="SPRODCODE" value="<%=SPRODCODE%>">
					<input type="hidden"  name="SEXPTYPE" value="<%=SEXPTYPE%>">
<%End if%>
			<input type="hidden"  name="spono" value="<%=Request("spono")%>">
					<input type="hidden"  name="sprno" value="<%=Request("sprno")%>">
                    <input type="submit" name="Submit" value="Save">
                    <input type="reset" name="Reset" value="Reset">
                    <br>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp; </td>
              </tr>
            </table>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
              | <a href="../workflow/workflow.asp">Workflow Manager</a></font></p>
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
Sub Update_Product(sUpdateMode)
	Dim  Rs, sSQL
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		sprodcode = Request.Form("sprodcode")
		sexptype = Request.Form("sexptype")
		svendcode = Request.Form("svendcode")
		sprodid = Request.Form("sprodid")
		stprodname = Request.Form("stprodname")
		seprodname = Request.Form("seprodname")
		smodel = Request.Form("smodel")
		sAcccode = Request.Form("sAcccode")
		stdesc = Request.Form("stdesc")
		sedesc = Request.Form("sedesc")
		nprice = Request("nprice")
		sCatID = Request.Form("sCatID")
		simage=Request.Form("simage")
		sempno = Session("sempno") 'Request.Form("sempno")
		sCurcode=Request("sCurrency")
		sUnitcode=Request("sUnit")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(sprodcode) = "") Then
		bError = True
		img_sprodcode = sErrImage
	End If

'	If (Trim(sexptype) = "") Then
'		bError = True
'		img_sexptype = sErrImage
'	End If

'	If (Trim(svendcode) = "") Then
'		bError = True
'		img_svendcode = sErrImage
'	End If
 if  Not IsNumeric(NPRICE) then  
 		bError = True
		img_nprice = sErrImage
	End If

	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  PProduct WHERE sprodcode='" & sprodcode &"'and sexptype='" & sexptype &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "pproduct", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("sprodcode") = sprodcode
		Rs("sexptype") = sexptype
		Rs("scatID") = scatID
		Rs("svendcode") = svendcode
		Rs("sprodid") = sprodid
		Rs("stprodname") = stprodname
		Rs("seprodname") = seprodname
		Rs("smodel") = smodel
		Rs("sAcccode") = sAcccode
		Rs("stdesc") = stdesc
		Rs("sedesc") = sedesc
		Rs("nprice") = FormatNumber(nprice, 4)
		Rs("simage") = sfilename
	   	Rs("dadddate") = date
		Rs("sempno") = sempno
		RS("sCurcode")=sCurcode
		RS("sUnitcode")=sUnitcode
	  Rs.Update	

	  Rs.Close
	  Set Rs = Nothing
'if sMode = "ADD" then
'     Response.Redirect "pproductdesc.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"&sCatid="&sCatID
'else

     Response.Redirect "pproductedt.asp?sprodcode="&sprodcode&"&sexptype="&sexptype&"sCatid="&sCatID&"sStatus=del"
'End if
     'Response.Redirect "javascript:history.back()"
	else

		Show_Form()
	End If	' if Error 
End Sub	
Sub List_PCatalog(sCatID)
Dim sSQL2,RS2,cSelect
		sSQL2= "Select * from pCatalog"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText
	IF NOT Rs2.EOF  Then
				Do While NOT Rs2.EOF
					If Rs2("sCatID") = sCatID Then
						cSelect = "selected"
					Else
						cSelect = ""
					End IF
			Response.Write "<option value= '"&Rs2("sCatID")& "'" & cSelect &">"& Rs2("sCatalog") &"</option>"& vbCrLf
			Rs2.MoveNext
		Loop
	End IF
	set Rs2=Nothing
End Sub

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
			Response.Write(Rs2("sCatdesc"))
			Rs2.MoveNext
		Loop
	End IF
End Sub
%>
