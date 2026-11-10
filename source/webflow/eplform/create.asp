<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->

<%
Dim i, rsOra, rsOra2, sqlOra
Dim sign, url
	Dim protocol
	If LCase(Request.ServerVariables("HTTPS")) = "on" Then
		protocol = "https://"
	Else
		protocol = "http://"
	End If

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

validateUser
'Cancle by User 
'Remove 26/07/2013
'if timeIsOK() then
'	response.write "<font color=#FF0000>Cannot Request form after 16:30</font>"
'	response.end
'end if
%>

<html>
<head>
<!meta http-equiv="refresh" content="60;url=index.asp">
<meta http-equiv="Content-Type" content="text/html; charset=tis-620">
<link href="../css/common.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">

<table><tr><td></td>
    <td><b>Create EPL Forms</b></td>
  </tr></table><%
sqlOra = "select " _
& "formMst.*, to_char(nNo) as nNo2 " _
& "from formMst " _
& "where " _
& "cStatus = '" & FORM_STS_SHOW & "' and vOrgNo like '070%' and nno <> '2' and cyear <> '14' " _
& "order by vName"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	response.write "<table cellpadding=1 cellspacing=0 border=0>"

	do while not rsOra.EOF
			if inStr(rsOra("vFormPage"), "?") = 0 then
				sign = "?"
			else
				sign = "&"
			end if
			
		'Display form name
		if not isNull(rsOra("vFormPage")) and session("uid") <> "" then
		%>
		<tr>
		  <td><img src="../images/orange_bullet.gif"></td>
		  <td>
		    <a href="<%=rsOra("vFormPage") & sign & "no=" & rsOra("nNo2") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear") & "&bp=" & "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & server.urlEncode("/create.asp?orgNo=") & rsOra("vOrgNo")%>"><%=rsOra("vName")%></a>
	<%	if (not isNull(rsOra("vDesc")) or rsOra("vDesc") <> "") then
			response.write " <font color=""gray"">(" & rsOra("vDesc") & ")</font>"
		end if	%>
			
		  </td>
		 </tr>
		<%
		else
			response.write "<li>" & rsOra("vName") '& "<br>"
		end if
		rsOra.moveNext
	loop
%>
    <tr>
		<td><img src="../images/orange_bullet.gif"></td>
		<td><a href="<%=protocol&"//webflow.mitsubishielevatorasia.co.th/form/eplform/coversheet/index.asp?"%>">Create Cover sheet</a></td>
	</tr>
<%	
else
	response.write "<tr><td>Form not found!</td></tr>"
end if

rsOra.close
response.write "</table>"
%>

</body> 
</html> 

