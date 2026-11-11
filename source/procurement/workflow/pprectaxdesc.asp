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
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim sprecno,sprno,drecdate,srecby,sreceiver,d_day,d_month,d_year
	Dim img_srecno,sbread,Item,sPono
	Dim skeyby
	Dim prectime , pplace
	sprecno=Request("sprecno")
	sMode=Request("Event")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")
	
' This routine retrieves the member by the memberID. If the event is not found,
' set the mode to an "ADD".
  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM planreceive WHERE id='" & sprecno & "'" 
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sprecno=RS("ID")
		sprno = Rs("sprno")
		sPono=Rs("sPono")
		drecdate=RS("precdate")
		d_day=day(drecdate)
		d_month=month(drecdate)
		d_year=year(drecdate)
		prectime = Rs("prectime")
		pplace = Rs("precplace")
		sbread="ReadOnly"
		skeyby = RS("skeyby")
	Else
		response.write("No Data")
	End If
%>	

<html>
<head>
<title>Edit Plan Receive :: Computer</title>
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
          <td width="96%" align="left" valign="top"><b><img src="../images/plus.gif" width="9" height="9"> 
            WORK FLOW MANAGER</b></td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" height="166">&nbsp;</td>
          <td width="96%" align="center" valign="top" height="166"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="92%" bgcolor="#DB4B37"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Plan Receive 
                  Management :</font></b></font> <font color="#FFFFFF" size="2"><%=sMode%> 
                  Plan Receive</font></b></font></td>
                <td width="8%" bgcolor="#8B2718" align="center"><img src="../images/iconhelp2.gif" width="30" height="26"></td>
              </tr>
              <tr bgcolor="#F8D7D1" align="center" valign="top"> 
                <td colspan="2"> 
                  <table width="98%" border="0" cellspacing="1" cellpadding="2">
                    <tr> 
                      <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/R 
                        No. </font></b></td>
                      <td width="77%"><a href="pprdesc.asp?sprno=<%=sprno%>"><%=sprno%> </a></td>
                    </tr>
                    <tr> 
                      <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">P/O
                        No. </font></b></td>
                      <td width="77%"><%=spono%> </td>
                    </tr>
                    <tr> 
                      <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receive 
                        Date </font></b></td>
                      <td width="77%"> <%=show_date(drecdate)%> </td>
                    </tr>
					   <tr> 
                      <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receive 
                       Time </font></b></td>
                      <td width="77%"><%=prectime%> </td>
                    </tr>
						   <tr> 
                      <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receive 
                       Place </font></b></td>
                      <td width="77%"><%=pplace%> </td>
                    </tr>
                    <tr> 
                      <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Plan Receive 
                        No. </font></b></td>
                      <td width="77%"><%=sprecno%> </td>
                    </tr>
                    <tr> 
                      <td width="23%" align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#CC3333">Planer</font></b></td>
                      <td width="77%">(<%=skeyby%>) <%Show_Empno(skeyby)%>  &lt;<font size="1"><%Show_Org(skeyby)%></font>&gt;</td>
                    </tr>				
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2" height="2">&nbsp; <% 
	If sMode="Show" Then
		Show_Item()
	Else
		Show_Item_receive()
	End if
%></td>
              </tr>
            </table>
            <br>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:history.back()">back</a> 
            | <a href="workflow.asp">WorkFlow Manager</a></font> </td>
          <td width="2%" height="166">&nbsp;</td>
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
<br></body>
</html> <%
	Sub Show_Item_Receive()
%> <%
    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color,Counter,chkreceive,chkno
Counter = 0
Color="#F8D7D1"
	sExptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  * FROM planreceiveitem  where ID='"&spRecno&"'" & " Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText


%> 
<form method="post" action="../workflow/Receive.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Item</font></b></font>
					</td>
                <td  width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Name</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Quantity</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">U/M</font></b></font>
					</td>
                <td width="8%" bgcolor="#8B2718">&nbsp;
					</td>
                <td width="9%" bgcolor="#8B2718" align="right">
						
              <DIV Valign="top"> </DIV>
					</td>
              </tr>
<% 
counter=0
Do While Not RS.EOF

'End if
Counter = Counter + 1
%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nItem")%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdname")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nQuality"), 3)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%>
						</td>
                <td width="1%" bgcolor="<%=Color%>" align="center" height="2" valign="top">

						</td>
                
            <td width="16%" bgcolor="<%=Color%>" align="left" height="2"> 
              <input type="radio" name='<%=Counter%>' value="2"> Receive 
						<input type="radio" name='<%=counter%>' value="1" Checked>
              No </td>
              </tr>
<%

Item = cInt(RS("nItem"))+1
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
RS.MoveNext
Loop

%>
<tr>
            <td colspan="6" height="2">&nbsp;</td>
            <td width="8%" bgcolor="#FFD553" align="center" height="2"></td>
</tr>
            </table>
<DIV Align="Center">
							<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno">
							<INPUT type="hidden" Value='<%=counter%>' name="counter">
							<INPUT type="hidden" Value='<%=Counter%>' name="Count_Item">

                     <input type="submit" name="Submit2" value="Receive">

</DIV>
          </td>
        </tr>
       </table>
</Form>

<%
End Sub


	Sub Show_Item()

    Dim  Rs,sSQL,Rs2,sSQL2
	Dim sScriptPath,sexptype,Color,Counter,chkreceive,chkno
Counter = 0
Color="#F8D7D1"
	sExptype = Session("sexptype")
	'scorpcode = session("gbcorpcode")
	' Get the name of the file

	sScriptPath = Request.ServerVariables("SCRIPT_NAME")

	Set Rs = Server.CreateObject("ADODB.RecordSet")       	

	sSQL = "SELECT  * FROM planreceiveitem  where ID='" & spRecno & "'" &_
	" Order By nItem"

	' Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

	'RS.PageSize = 5
'response.write RS2("MaxItem")	
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
%>
                  <form method="post" action="../workflow/Receive.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0">

        <tr bgcolor="#FFD553"> 

          <td width="96%" align="left" valign="top" bgcolor="#FFD553"> 
            
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr bgcolor="#DB4B37"> 
                <td width="5%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Item</font></b></font>
					</td>
                <td  width="25%" bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Name</font></b></font>
					</td>
                <td  width="10%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">Quantity</font></b></font>
					</td>
                <td  width="5%"  bgcolor="#8B2718"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666">
						<b><font color="#FFFFFF" size="2">U/M</font></b></font>
					</td>

<!--                <td width="8%" bgcolor="#8B2718">&nbsp;<img src="../images/dot.gif">
					</td>
                <td width="9%" bgcolor="#8B2718" align="right">
					<img src="../images/dot.gif">
					</td> -->
              </tr>
<% 
counter=0
if Not RS.EOF then
Do While Not RS.EOF
'if RS("cStatus") <> 0  Then
'End if
Counter = Counter + 1
%>

              <tr bgcolor="<%=Color%>"> 
                <td  bgcolor="<%=Color%>" height="2"><%=RS("nItem")%>&nbsp;
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sProdname")%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=FormatNumber(RS("nQuality"),3)%>
						</td>
                <td  bgcolor="<%=Color%>" height="2"><%=RS("sUnitcode")%>
						</td>
<!--                <td width="1%" bgcolor="<%=Color%>" align="center" height="2" valign="top">
					<img src="../images/dot.gif">
						</td>
                <td width="16%" bgcolor="<%=Color%>" align="center" height="2"><img src="../images/dot.gif"> -->
<%
'			If RS("cStatus")="2" then
%><!-- <img src="../images/istepapp.gif" width=20 height=19>
<%'Else %>
<img src="../images/istepappw.gif" width=20 height=19> -->
<%			'End if
%>
					</td>
              </tr>
<%

Item = cInt(RS("nItem"))+1
	if Color="#F8D7D1" then
			Color="#F1B3A7"
	else
			Color="#F8D7D1"
	end if
'End if
RS.MoveNext
Loop
	
else
	response.write "<center><a href ='precdel.asp?spRecno="&spRecno&"&sprno="&sprno&"&spono="&spono&"'> Delete Receive </a></center>"
End if
Rs.Close
%>
</tr>
            </table>
<DIV Align="Center">
							<INPUT type="hidden" Value='<%=sPrno%>' name="sPrno">
							<INPUT type="hidden" Value='<%=counter%>' name="counter">
							<INPUT type="hidden" Value='<%=Counter%>' name="Count_Item">

 </DIV>
          </td>
        </tr>
       </table>
</Form>
<%
'Response.write("sPrno = " &sPrno)
'Response.write("<BR>sExptype = "&sExptype)
'Response.write("<BR>sSQL = "&sSQL)
'Response.write("<BR>nItem = "&Item)
'Response.write("<BR>Counter = "&Counter)

End Sub
%>

<%
Sub Show_Org(sEmpno)
	Dim RS2,sSQL2,RS3,sSQL3
	sSQL2="Select * from vAemployee where sEmpno='" & sEmpno & "'"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")      
		Set Rs3= Server.CreateObject("ADODB.RecordSet")      
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not RS2.EOF then
		sSQL3="Select * from pSection where sSeccode='" & RS2("sSeccode") & "'"
		 	Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
			If Not RS3.Eof then
				Response.write  RS3("sSec")
			End if
		RS3.Close
			sSQL3="Select * from pDepartment where sDepcode='" & RS2("sDepcode") & "'"
		 	Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
			If Not RS3.Eof then
				Response.write  "/" & RS3("sDept")
			End if
		RS3.Close
		Set RS3=Nothing
		RS2.Close
		Set RS2=Nothing
	End if
End Sub
%>