<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file ="../list.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim svendcode,stname,sename,saddress1,saddress2,stelno,sfax,scontact,sposition,stnote,dadddate,sempno,sTermcode,sTermPayment,sStdCur, sType
	Dim sSTname,sSEname,sEmail,sCity,sState,sPost,sSearchKey,sPurAddress1,sPurAddress2,sPurCity,sPurState,sPurPost,sPurContact,sPurPhone,spurEmail
	Dim img_sVendcode	,img_stname,img_sename,img_staddress,img_seaddress,img_stelno
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year,sbread,status,msg
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
		'Show_Form sMode , svendcode
		Show_Form sMode
	Else
		Update_Vender(Request.Form("EVENT"))
	End If		
	
	'response.write "Event = "&Request.Form("EVENT")
	
%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
Sub Get_Vender(svendcode)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
 'Response.write "MODE = "&sMode
  sSQL = "SELECT * FROM Pvender WHERE svendcode='" & svendcode & "'"
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		svendcode = Rs("sVendcode")
		stname = Rs("sTname")
		sename = Rs("sEname")
		saddress1 = Rs("sEAddress")
		saddress2 = Rs("sTAddress")
		stelno = Rs("sTelno")
		sfax = Rs("sFax")
		scontact = Rs("sContact")
		sposition = Rs("sPosition")
		sCity=Rs("sCity")
		sState=Rs("sState")
		sPost=Rs("sPost")
		sSearchKey=Rs("sSearchkey")
		sPurAddress1=Rs("sPurAddr1")
		sPurAddress2=Rs("sPurAddr2")
		sPurCity=Rs("sPurCity")
		sPurState=Rs("sPurState")
		sPurPost=Rs("sPurPost")
		sPurContact=Rs("sPurContact")
		sPurPhone=Rs("sPurPhone")
		stnote = Rs("sTnote")
		dadddate = Rs("dAdddate")
		d_day=day(dadddate)
		d_month=month(dadddate)
		d_year=year(dadddate)
		sempno = RS("sempno")
		sSTname=RS("sSTname")
		sSEname=Rs("sSEname")
		sEmail=Rs("sEmail")
		sPurEmail=Rs("sPurEmail")
		sTermCode=RS("sTermCode")
		sTermPayment=RS("sTermpayment")
		sStdCur=RS("sStdCur")
		sbread="ReadOnly"
	Else
		sMode = "ADD"
	End If
	'response.write "sMode222222 = "&sMode
End Sub	

%>	

<%
' This routines presents the form to the user. Depending on the mode, the
' form fields are filled with data or blank.
'Sub Show_Form(mode,vend)
Sub Show_Form(mode)
	' Set the title of the page depending on if an error occurred during
	' form submission or not
	
	'response.write "INSHOWFORM vend 1 = "&svendcode
	if sMode = "" then
		sMode = mode
	end if
	
	'response.write "smode = "&sMode
	'if svendcode = "" then
	'	svendcode = vend
	'end if
	
%>
<html>
<head>
<title>Edit Vendors :: Computer</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
<SCRIPT LANGUAGE="JavaScript">
<!--
function Form_Load(){
	document.frmName.SVENDCODE.focus();
}
-->
</SCRIPT>

</head>

<%IF sMode="EDIT" then 
	Response.write "<body bgcolor='#FFFFFF' leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
Else %>
	<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="Form_Load()" >
<%End IF %>
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
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  Vendors </font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="26" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>" name="frmName">
                    <table width="98%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td width="21%" align="Right"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Vendor 
                          Code : </font></b></td>
                        <td width="79%"> 
                          <%If sMode="EDIT" then
								Response.write sVendcode
								
							%>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="scm" id="scm" value="1"  <% if CheckSCM(sVendcode) = "1" then %> checked <% end if %> />&nbsp;&nbsp;Update scm vendor
							<%Else%>
                          <input type="text" name="SVENDCODE" size="12" maxlength="12" value="<%=svendcode%>" <%=sbread%>>
						  <input type="checkbox" name="scm" id="scm" value="1" />&nbsp;&nbsp;Add to be scm vendor
                          <%=img_svendcode%> 
                          <%End If%>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor Code Data 
                          (Character) :Not more than 12 Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">English 
                          Name : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SENAME" size="50" maxlength="90" value="<%=sename%>"><%=img_sename%> 
                        </td>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Name (Character,English 
                          Language) :Not more than 90 Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">English 
                          Short Name : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SSENAME" size="31" maxlength="30" value="<%=ssename%>"> 
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Short Name 
                          (Character,English Language) :Not more than 30 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"></font> 
                          <font color="#CC0000"><b>*</b></font> <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Thai 
                          Name : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="STNAME" size="50" maxlength="90" value="<%=stname%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Name (Character,Thai 
                          Language) :Not more than 90 Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Thai 
                          Short Name : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SSTNAME" size="31" maxlength="30" value="<%=sstname%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Short Name 
                          (Character,Thai Language) :Not more than 30 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right" valign="top"><b><font color="#CC0000"><b>*</b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Address1 : </font></b></td>
                        <td width="79%"> 
                          <textarea name="SADDRESS1" cols="45" rows="3"><%=saddress1%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Address 
                          (Character,Thai Language) :Not more than 120 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Address 2 : </font></b></td>
                        <td width="79%"> 
                          <textarea name="SADDRESS2" cols="45" rows="3"><%=saddress2%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Address 
                          (Character,English Language) :Not more than 30 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">City. 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SCITY" size="30" maxlength="60" value="<%=sCity%>">
                          <font color="#CC3333">(Character) :Not more than 60 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">State. 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SSTATE" size="30" maxlength="60" value="<%=sState%>">
                          <font color="#CC3333">(Character) :Not more than 60 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Post. 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPOST" size="6" maxlength="5" value="<%=sPost%>">
                          <font color="#CC3333">(Character) :Not more than 5Characters 
                          .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"><b></b></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Tel. 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="STELNO" size="15" maxlength="30" value="<%=stelno%>">
                          <font color="#CC3333">(Character) :Not more than 30 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Fax. 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SFAX" size="15" maxlength="30" value="<%=sfax%>">
                          <font color="#CC3333">(Character) :Not more than 30 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Email 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SEMAIL" size="30" maxlength="30" value="<%=sEmail%>">
                          <font color="#CC3333">(Character) :Not more than 30 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Contact 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SCONTACT" size="45" maxlength="90" value="<%=scontact%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Person contact (Character) 
                          :Not more than 60 Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Position 
                          : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPOSITION" size="30" maxlength="30" value="<%=sposition%>">
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Position of Person 
                          contact (Character) :Not more than 30 Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">Search 
                          Key : </font></b></td>
                        <td width="79%"> 
                          <textarea name="sSearchkey" cols="45"><%=sSearchkey%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Search key (Character) 
                          :Not more than 120 Characters .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right" valign="top"><b><font color="#CC0000"></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Pur Address1 : </font></b></td>
                        <td width="79%"> 
                          <textarea name="SPURADDRESS1" cols="45" rows="3"><%=spuraddress1%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Address 
                          (Character,Thai Language) :Not more than 120 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          Pur Address 2 : </font></b></td>
                        <td width="79%"> 
                          <textarea name="SPURADDRESS2" cols="45" rows="3"><%=spuraddress2%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Vendor's Address 
                          (Character,English Language) :Not more than 30 Characters 
                          .</font></td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Pur 
                          City. : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPURCITY" size="30" maxlength="60" value="<%=sPurCity%>">
                          <font color="#CC3333">(Character) :Not more than 60 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"><b> 
                          </b></font> <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Pur 
                          State. : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPURSTATE" size="30" maxlength="60" value="<%=sPurState%>">
                          <font color="#CC3333">(Character) :Not more than 60 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Pur 
                          Post. : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPURPOST" size="6" maxlength="5" value="<%=sPurPost%>">
                          <font color="#CC3333">(Character) :Not more than 5Characters 
                          .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font color="#CC0000"></font> 
                          <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Pur 
                          Tel. : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPURPHONE" size="15" maxlength="30" value="<%=sPurphone%>">
                          <font color="#CC3333">(Character) :Not more than 30 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Pur 
                          Email : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="SPUREMAIL" size="15" maxlength="30" value="<%=sPurEmail%>">
                          <font color="#CC3333">(Character) :Not more than 30 
                          Characters .</font> </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Term Code : </font></b></td>
                        <td width="79%"> 
                          <Select name="stermcode">
						  <%List_termcode(stermcode)%>
						  </Select>
                          <font color="#CC3333">(Character) :Not more than 120 
                          Characters .</font> </td>
                      </tr>
<!--                      <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Term Payment : </font></b></td>
                        <td width="79%"> 
                          <input type="text" name="sTermPayment" size="45" maxlength="255" value="<%=sTermPayment%>">
                          <font color="#CC3333">(Character) :Not more than 255
                          Characters .</font> </td>
                      </tr> -->
                     <tr> 
                        <td width="21%"  align="Right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Std Currency : </font></b></td>
                        <td width="79%"> 
                           <select name="sStdCur" size=1>
                            <%List_Currency(sStdCur)%> 
                          </select>
                           </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Note 
                          : </font></b></td>
                        <td width="79%"> 
                          <textarea name="STNOTE" cols="45" rows="3"><%=stnote%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td width="21%"  align="Right">&nbsp;</td>
                        <td width="79%"><font color="#CC3333">Any more Data :Not 
                          more than 120 Characters .</font></td>
                      </tr>
					  <%If sMode="EDIT" then %>
					   <tr> 
                        <td width="21%"  align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Disable</font></b></td>
                        <td width="79%"><input type="checkbox" name="mark" value="F" <% if InStr(sename, "(DO NOT USE)") <> 0 then Response.write "checked" end if %> >
                        </td>
                      </tr>
					  
					  <!--ADD CHECK BOX FOR RE-SENT DATA TO SCM -->
					  <tr> 
                        <td width="21%"  align="Right" valign="top"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Resent to SCM</font></b></td>
                        <td width="79%"><input type="checkbox" name="resenttoscm" value="1"/>
                        </td>
                      </tr>
					  <% END IF %>
                      <tr> 
                        <td width="21%"  align="Right" valign="top">&nbsp;</td>
                        <td width="79%"><font color="#CC0000"><b>*</b></font> 
                          Required Field.</td>
                      </tr>
                      <!--                      <tr> 
                        <td width="19%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#CC3333">AddDate</font></td>
                        <td width="81%"><font face="MS Sans Serif" size="2">วันที่</font> 
                          <select name="dadddate_day">
                            <% '   if d_day="" then d_day=day(date)
												'		Create_daylist(d_day)  
								%> 
                          </select>
                          <font size="2" face="MS Sans Serif">เดือน</font> 
                          <select name="dadddate_month">
                            <%' if d_month="" then d_month=month(date)
											'					Create_MonthList(d_month)
								%> 
                          </select>
                          <font size="2" face="MS Sans Serif">พ.ศ.</font>. <% 'if d_year="" then d_year=year(date) %> 
                          <input type="text" name="dadddate_year" size="8" maxlength="4" value="<%=d_year%>">
							 </td>
                      </tr>
                      <tr> 
                        <td width="19%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Empno</font></td>
                        <td width="81%"> 
                          <input type="text" name="SEMPNO" size="12" maxlength="12" value="<%=sempno%>">
                        </td>
                      </tr> --> 
                    </table>
                    <br>
                    <input type=HIDDEN name=EVENT value="<%=sMode%>">
                    <%
					'Kanitha Modify sMode <> "ADD" show hidden
					'If sMode<>"ADD" and bError=false then
					If sMode<>"ADD" then
					%> 
                    <input type="HIDDEN" name="sVendcode" value="<%=sVendcode%>">
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
          <td width="2%" align="left" valign="bottom"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="10" height="12"></td>
          <td width="2%" align="right" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
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



Function CheckSCM(vendcode)

	Dim Rs2 , sSQL 
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL = "select * from scmvender where svendcode = '"&vendcode&"'"
	Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If NOT Rs2.EOF Then
		CheckSCM = 1
	ELSE
		CheckSCM = 0
	END IF


End Function

Sub UpdateSCM(vendcode)
	Dim Rs2 , sSQL 
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL = "select * from scmvender where svendcode = '"&vendcode&"'"
	Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Rs2.EOF Then
		Rs2.close
		sSQL = "insert into scmvender(svendcode) values('"&vendcode&"')"
		Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText

	END IF
End Sub

Sub KeepLog(status,comment,vendcode,vendname)
	Dim Rs2 , sSQL , id
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	sSQL = "select (count(logid)) as id from logvendor"
	Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs2.EOF Then
		If Rs2("id") = "0" Then
			id =  1
		ELSE
			Rs2.close
			sSQL = "select max(logid) as id from logvendor"
			Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			id = cint(Rs2("id")) + 1
		END IF
		'response.write id
		Rs2.close
	'	sSQL = "insert into logvendor values('"&id&"',to_date(to_char(sysdate,'dd/mm/yyyy'),'dd/mm/yyyy'),'"&Session("sEmpno")&"','"&Session("sName")&"','"&status&"','"&comment&"','"&vendcode&"','"&vendname&"')"
	'	response.write sSQL
	'	Rs2.Open sSQL, conn
		Rs2.Open "logvendor", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		Rs2.AddNew
		Rs2("logid") = id
		Rs2("logdate") = Now
		Rs2("USERID") = Session("sEmpno")
		Rs2("USERNAME") = Session("sName")
		Rs2("STATUS") = status
		Rs2("VCOMMENT") = comment
		Rs2("SVENDCODE") = vendcode
		Rs2("SENAME") = vendname	
		Rs2.Update	
		'response.end
	END IF
	Rs2.Close
	'Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	'response.write sSQL
	'if Not Rs.EOF Then
	'	response.write "IFFFFFFFFFFFFF"
	'END IF
	'if Rs2.EOF then
	'	response.write "IFffffffffffffff"
	'else
	'	response.writ "ELSEeeeeeeeeeee"
	'end if
End Sub


Sub Update_Vender(sUpdateMode)
	Dim sErrImage, Rs, sSQL , Rs2 , tempsname
	bError = False
  	sErrImage = "<IMG SRC='images/icerror.gif'>"
	'sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
		'response.write "test = "&Request.Form("ark")
		svendcode = Request.Form("sVendcode")
		stname = Ucase(Request.Form("sTname"))
		sename = Ucase(Request.Form("sEname"))
		saddress1 = Request.Form("saddress1")
		saddress2 = Request.Form("saddress2")
		stelno = Request.Form("sTelno")
		sfax = Request.Form("sFax")
		scontact = Request.Form("sContact")
		sposition = Request.Form("sPosition")
		sCity=Request.Form("sCity")
		sState=Request.Form("sState")
		sPost=Request.Form("sPost")
		sSearchkey=Ucase(Request.Form("sSearchkey"))
		sPurAddress1=Request.Form("spurAddress1")
		sPurAddress2=Request.Form("sPurAddress2")
		sPurCity=Request.Form("sPurCity")
		sPurState=Request.Form("sPurState")
		sPurPost=Request.Form("sPurPost")
		sPurContact=Request.Form("sPurContact")
		sPurphone=Request.Form("sPurPhone")
		stnote = Request.Form("sTnote")
		sSTname=Request.Form("sSTname")
		sSEname=Request.Form("sSEname")
		sEmail=Request.Form("sEmail")
		spurEmail=Request.Form("sPurEmail")
		sTermCode=Request.Form("sTermcode")
		sTermPayment=Request.Form("sTermPayment")
		sStdCur=Request.Form("sStdCur")
		dadddate = date'Request.Form("dadddate_month")&"/"&Request.Form("dadddate_day")&"/"&Request.Form("dadddate_year")
		sempno = Session("sempno")

	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

	If (Trim(svendcode) = "") Then
		bError = True
		img_svendcode = sErrImage
	End If

	If (Trim(stname) = "") Then
		bError = True
		img_stname = sErrImage
	End If

	If (Trim(sename) = "") Then
		bError = True
		img_sename = sErrImage
	End If

	If (Trim(saddress1) = "") Then
		bError = True
		img_staddress = sErrImage
	End If
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	
	'Response.write "updatemode = "&sUpdateMode
	IF sUpdateMode = "ADD" Then
		If (Trim(svendcode) <> "" ) Then
			sSQL = "SELECT * FROM  PVENDER WHERE svendcode='" & Trim(svendcode) &"'"
			Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
			if Not Rs2.EOF then
					bError = True
					img_svendcode = "Duplicate Vendor Code "&sErrImage
			end if
			Rs2.close
		End if
		
		'If (Trim(sename) <> "" ) Then
		'	sSQL = "SELECT * FROM  PVENDER WHERE REPLACE(REPLACE(sename,' ',''),'(DONOTUSE)','')='" &Replace(Replace(sename," ",""),"(DONOTUSE)","")&"' "
			'response.write sSQL
		'	Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
		'	if Not Rs2.EOF then
		'			bError = True
		'			img_sename = "Duplicate Vendor "&Rs2("SENAME")&""&sErrImage
		'	end if
		'	Rs2.close
		'End if
	End IF
	'IF sUpdateMode = "EDIT" Then	
	'	If (Trim(sename) <> "" ) Then
	'		sSQL = "SELECT * FROM  PVENDER WHERE REPLACE(REPLACE(sename,' ',''),'(DONOTUSE)','')='" &Replace(Replace(sename," ",""),"(DONOTUSE)","")&"' and svendcode <> '"&svendcode&"'"
			'response.write sSQL
	'		Rs2.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	'		if Not Rs2.EOF then
	'				bError = True
	'				img_sename = "Duplicate Vendor "&Rs2("SENAME")&""&sErrImage
	'		end if
	'		Rs2.close
	'	End if
	'End IF
	

'	If Trim(stelno) = "" Then
'		bError = True
'		img_stelno = sErrImage
'	End IF

	If bError = False Then
	  ' Create a recordset
	

		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  PVENDER WHERE svendcode='" & svendcode &"'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  'response.write "ELSE"
		  Rs.Open "pvender", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
		Rs("svendcode") = Trim(svendcode)
		if Request.Form("mark") = "F" then
			if InStr(Rs("sename"), "(DO NOT USE)") = 0 then
				Rs("sename") = "(DO NOT USE)"&sename
			else
				Rs("sename") = sename
			end if
		else
			if InStr(Rs("sename"), "(DO NOT USE)") <> 0 then
				Rs("sename") = Replace(sename,"(DO NOT USE)","")
			else
				Rs("sename") = sename
			end if
		end if 
		
		tempsname = Rs("sename")
		Rs("stname") = stname
		Rs("sEaddress") = saddress1
		Rs("sTaddress") = saddress2
		Rs("stelno") = stelno
		Rs("sfax") = sfax
		Rs("scontact") = scontact
		Rs("sposition") = sposition
		RS("sCity")=sCity
		RS("sState")=sState
		RS("sPost")=sPost
		RS("sSearchKey")=sSearchKey
		RS("sPurAddr1")=sPurAddress1
		RS("sPurAddr2")=sPurAddress2
		RS("sPurCity")=sPurCity
		RS("sPurState")=sPurState
		RS("sPurPost")=sPurPost
		RS("sPurContact")=sPurContact
		RS("sPurPhone")=sPurPhone
		Rs("stnote") = stnote
	   	Rs("dadddate") = dadddate
		Rs("sempno") = sempno
		RS("sSTname")=sSTname
		RS("sSEname")=sSEname
		RS("sEmail")=sEmail
		Rs("sPurEmail")=sPurEmail
		RS("sTermCode")=sTermcode
		RS("sTermPayment")=sTermPayment
		RS("sStdCur")=sStdCur
	  Rs.Update	
	  Rs.Close
	  
		'Add data export to SCM
		if CheckSCM(sVendcode) = "1" then 
			sType = 1
		else
			sType = 2
		end if
		
		IF sUpdateMode = "ADD" Or Request.Form("resenttoscm") = "1" Then
		sSQL = "insert into EXPVENDOR values('"& svendcode & "|" & sename & "|" & saddress1 & "|" & saddress2 & "||" & sCity & "|" & sPost & "||" & sType  & "|" & Session("sempno") & "','" & Session("sempno") &"')"
		'Response.write sSQL
		'response.end

		Rs.Open sSQL, conn
		end if
		'Rs.Close
	  if sUpdateMode = "EDIT" Then
				response.write "scm=>>>"&Request.Form("scm") 
				'response.end
				if Request.Form("scm") = "1" then
					UPDATESCM svendcode
				end if
			    KeepLog "U","Update Vendor", svendcode ,tempsname
				msg = " To: PUR(PIC), PUR SEM, PUR DDEM, PUR DEM, PUR DIM" _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"    Input date: " & Now _
					  & chr(13) & chr(10) _
					  & "    Vendor status are transferred  from Workflow systerm .  Please see as below." _
					  & chr(13) & chr(10) _
					  &"    - 'Update' , Vendor code : "&Trim(svendcode)&" , Vendor name : "&tempsname&" , Input by "&Session("sName") _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"  Best Regards," _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"  Auto Send mail System. "
				'msg = "Vendor Code "&svendcode&" is Enable by "&Session("sName")
				'sendMail MAIL_SERVER, SYSTEM_NAME, SYSTEM_ADMIN_MAIL, "kanittha@MitsubishiElevatorAsia.co.th", "Vendor Management", msg  
				'sendMail MAIL_SERVER, SYSTEM_NAME, SYSTEM_ADMIN_MAIL, "PURALERTMAIL@MitsubishiElevatorAsia.co.th", "Vendor Management", msg  
	  else
				KeepLog "A","ADD Vendor", svendcode ,tempsname
				msg = " To: PUR(PIC), PUR SEM, PUR DDEM, PUR DEM, PUR DIM" _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"    Input date: " & Now _
					  & chr(13) & chr(10) _
					  & "    Vendor status are transferred  from Workflow systerm .  Please see as below." _
					  & chr(13) & chr(10) _
					  &"    - 'ADD' , Vendor code : "&Trim(svendcode)&" , Vendor name : "&tempsname&" , Input by "&Session("sName") _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"  Best Regards," _
					  & chr(13) & chr(10) _
					  & chr(13) & chr(10) _
					  &"  Auto Send mail System. "
				'msg = "Vendor Code "&svendcode&" is Enable by "&Session("sName")
				'sendMail MAIL_SERVER, SYSTEM_NAME, SYSTEM_ADMIN_MAIL, "kanittha@MitsubishiElevatorAsia.co.th", "Vendor Management", msg  	
				'sendMail MAIL_SERVER, SYSTEM_NAME, SYSTEM_ADMIN_MAIL, "PURALERTMAIL@MitsubishiElevatorAsia.co.th", "Vendor Management", msg  
	  end if	
	  Set Rs = Nothing
	status="del"
	'response.write bError&"IFFFFFFFFFFFFFF"
    Response.Redirect "../datamanager/pvender.asp?status=" & status

	else
		'response.write bError&"ELSEEEEEEEEEEEEE<br>"
		'response.write "vendor = "&svendcode
		'Show_Form sUpdateMode , svendcode
		Show_Form sUpdateMode
	End If	' if Error 
End Sub	

%>