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
on Error resume Next
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sprno,sexptype,dprdate,srequester,skey,cprtype,nstep,capprove
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Prno
		Item=1
	'Dim sKey
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
		if Request("sPrno") <> ""  then
				session("sPrno")=Request("sPrno")
				sprno=session("sprno")
				Get_Ppr(sPrno)
				Show_Form()
	End If		

%>

 <%
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".

Sub Get_Ppr(sprno)
	Dim sSQL, Rs
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM ppr WHERE sprno='" & sprno & "'" &_
"order by sprno"
   Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sprno=RS("sprno")
		session("sexptype")=RS("sexptype")
		dprdate=RS("dprdate")
		d_day=Day(dprdate)
		d_month=Month(dprdate)
		d_year=Year(dprdate)+543
		srequester=RS("srequester")
		skey=RS("skey")
		cprtype=RS("cprtype")
		nstep=RS("nstep")
		capprove=RS("capprove")
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
<title>PURCHASE REQUISITION</title>
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
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><img src="../images/y01.jpg" width="12" height="12"></font></td>
          <td width="96%" height="2" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">...</font></b></font></td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" height="113">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="113"> 
            <table width="100%" border="0" cellspacing="1" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2"><%=sMode%> 
                  PURCHASE REQUISITION</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <form method="post" action="<%=sScriptPath%>">
                    <table width="93%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">No. 
                          </font>&nbsp;<%=sprno%><font color="#FFCCCC">..</font></td>
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Expense 
                          Type </font>&nbsp;<% Show_Pexptype(session("sexptype")) %><font color="#FFCCCC">...</font> 
                        </td>
                      </tr>
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Date&nbsp;</font><%response.write Show_date(date)%><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font></td>
                        <td width="19%">&nbsp; </td>
                        <td width="19%"><font color="#FFCCCC">...</font></td>
                        <td width="26%"><font color="#FFCCCC">...</font> </td>
                      </tr>
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Request 
                          By &nbsp;</font><%=sRequester%><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333"> 
                          </font></td>
                        <td colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Due 
                          Date </font><%response.write Show_date(date)%></td>
                        <td width="26%"><font color="#FFCCCC">...</font> </td>
                      </tr>
                      <tr> 
                        <td colspan="3"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Key 
                          By </font>&nbsp;<%=sKey%></td>
                        <td width="19%"><font color="#FFCCCC">...</font></td>
                        <td width="19%"><font color="#FFCCCC">...</font> </td>
                        <td width="26%"><font color="#FFCCCC">...</font> </td>
                      </tr>
                    </table>
                  </form>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2">&nbsp;<% Show_Item(session("sexptype"))%></td>
              </tr>
            </table>
            <p><br>
              <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()"><font size="1">Back</font></a> 
              <font size="1">|</font> <a href="../datamanager/datamanager.asp"><font size="1">Data 
              Manager</font></a></font><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><br>
              </font></p>
          </td>
          <td width="2%" height="113">&nbsp;</td>
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
	Sub Show_Item(sexptype)
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,Color
Color="#F8D7D1"
	'sexptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  * FROM ppritem  where sprno='" & sPrno & "'" &_
	"Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

	'RS.PageSize = 5
'response.write RS2("MaxItem")	
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">

              <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Item</font></b></font></td>
                <td  width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Name</font></b></font></td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Quantity</font></b></font></td>
                <td  width="10%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Price</font></b></font></td>
                <td  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Use For</font></b></font></td>
                <td width="8%" bgcolor="#8B2718">&nbsp;</td>
                <td width="9%" bgcolor="#8B2718" align="right" valign="top"><img src="../Pexpenstype/images/y05.jpg" width="12" height="12"></td>
              </tr>
<% Do While Not RS.EOF %>
              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nitem")%>&nbsp;</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sprodname")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nQuality")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nPrice")%></td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUseFor")%></td>
                <td width="8%" bgcolor="<%=Color%>" align="center" height="2" valign="top"><a href="/amec/ppritem/ppritemedt.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>&sexptype=<%=sExptype%>"><img src="../images/iconedit2.gif" width="30" height="26" border="0" alt="EDIT"></a></td>
                <td width="9%" bgcolor="<%=Color%>" align="center" height="2"><a href="/amec/workflow/ppritemdel.asp?nitem=<%=RS("nitem")%>&sprno=<%=RS("sprno")%>"><img src="../images/icondelete2.gif" width="30" height="26" alt="DELETE" border="0"> 
                </td>
              </tr>
<%
if Color="#F8D7D1" then
		Color="#F1B3A7"
else
		Color="#F8D7D1"
end if
Item=Cint(RS("nItem"))+1
 RS.MoveNext
Loop
%>
<tr>
<td colspan="6">&nbsp;</td>
<td   width="8%" bgcolor="#FFD553" align="center" height="8"><a href="/amec/workflow/ppritemedt.asp?nItem=<%=Item%>&sPrno=<%=sPrno%>"><img src="../images/iconadd2.gif" width="30" height="26" alt="ADD" border="0"></a></td>
</tr>
            </table>

          </td>

        </tr>
 
      </table>

<%
End Sub

Sub Update_Ppr(sUpdateMode)
	Dim sErrImage, Rs, sSQL,dprdate
	bError = False
  	sErrImage = "<IMG SRC='../images/icerror.gif'>"
	'Response.Write "Input"
	' Get the form data into variables

	' Check to see if Start date is a valid date
'sprno=Request("sprno")
'Session("sexptype")=Request("sexptype")
'dprdate = Request.Form("dprdate_month")&"/"&Request.Form("dprdate_day")&"/"&Request.Form("dprdate_year")
dprdate=Date
'Session("srequester")=Request("srequester")
'skey=Request("skey")
cprtype=Request("cprtype")
nstep=Request("nstep")
capprove=Request("capprove")
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check if the End date is a valid date provided it is not empty

'	If (Trim(sprno) = "") Then
'		bError = True
'		img_sprno = sErrImage
'	End If
'	If (Trim(nstep) = "") Then
'		bError = True
'		img_step = sErrImage
'	End If
'	If (Trim(Request("dprdate_year")) = "") Then
'		bError = True
'		img_dprdate_year = sErrImage
'	End If
	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  ppr WHERE sprno='" & sprno &"'" &_
			"and sexptype='" & sexptype & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "ppr", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If
		'Response.Write		 sSQL
		' Populate the RecordSet fields	
		'dadddate=dadddate_month & "/" & dadddate_day & "/" & dadddate_year
if Prno="" then Prno=sPrno
		Rs("sprno") = Session("sPrno")
		Rs("sexptype") = Session("sexptype")
		Rs("dprdate") = dprdate
		Rs("srequester") = Session("srequester")
		Rs("skey") =Session("skey")
		Rs("cprtype") =cprtype
		Rs("nstep") =nstep
		Rs("capprove") =capprove
	  Rs.Update	
	  Rs.Close
	  Set Rs = Nothing
     Response.Redirect "ppr.asp?sexptype=" & sexptype & "&sKey=" & sKey & "&sprno=" & sPrno 
	else
		Show_Form()
	End If	' if Error 
End Sub	
%>