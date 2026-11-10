<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->

<%
Dim i, rsOra, rsOra2, sqlOra
Dim sign, url

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

validateUser
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<!meta http-equiv="refresh" content="60;url=index.asp">
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link href="../css/common.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">
<table>
  <tr>
	<td></td>
    <td><b>Create FE Electronic Forms</b></td>
  </tr>
</table>
<%

sqlOra = "select " _
& "formMst.*, to_char(nNo) as nNo2 " _
& "from formMst " _
& "where " _
& "cStatus = '" & FORM_STS_SHOW & "' and vOrgNo like '0510%' " _
& "order by vName"
'response.write sqlOra
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
		response.write "<tr><td><img src=""../images/orange_bullet.gif""></td><td>"

		if not isNull(rsOra("vFormPage")) and session("uid") <> "" and (getAuth(rsOra("nNo2"), rsOra("vOrgNo"), rsOra("cYear"), session("uid"), FORM_AUTH_ADMIN) or getAuth(rsOra("nNo2"), rsOra("vOrgNo"), rsOra("cYear"), session("uid"), FORM_KEY_OPR)) then
		%><a href="<%=rsOra("vFormPage") & sign & "no=" & rsOra("nNo2") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear") & "&bp=" & "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/" & MAT_SUB_SYSTEM_FOLDER & "/create.asp"%>"><%=rsOra("vName")%></a><%
		else
			response.write rsOra("vName")
		end if

		if not isNull(rsOra("vDesc")) or rsOra("vDesc") <> "" then
			response.write " <font color=""gray"">(" & rsOra("vDesc") & ")</font>"
		end if

		response.write "</td></tr>"
		rsOra.moveNext
	loop
else
	response.write "<tr><td>Form not found!</td></tr>"
end if

rsOra.close
%>
</table>

<table>
  <tr><td></td></tr>
  <tr><td><b>Working in factory master Management</b></td></tr>
  <tr><td><a href="<%="http://webflow/form/fe/index.asp?menu=subcon"%>"><img src="../images/orange_bullet.gif" border=0>Sub-contractor master</a></td></tr>
  <tr><td><a href="<%="http://webflow/form/fe/index.asp?menu=subcon_worker"%>"><img src="../images/orange_bullet.gif" border=0>Sub-contractor worker master</a></td></tr>
  <tr><td><a href="<%="http://webflow/form/fe/index.asp?menu=gas"%>"><img src="../images/orange_bullet.gif" border=0>Gas master</a></td></tr>
  <tr><td><a href="<%="http://webflow/form/fe/index.asp?menu=gas_sub"%>"><img src="../images/orange_bullet.gif" border=0>Gas supervisor</a></td></tr>
  <tr><td><a href="#"><img src="../images/orange_bullet.gif" border=0>Hot work scope master</a></td></tr>
  <tr><td><a href="<%="http://webflow/form/fe/index.asp?menu=item_mstd"%>"><img src="../images/orange_bullet.gif" border=0>Checking item master</a></td></tr>
  <tr><td><a href=#"><img src="../images/orange_bullet.gif" border=0>Machine installation and Construction inspection sheet</a></td></tr>
</table>

<table>
  <tr><td></td></tr>
  <tr><td><b>Report</b></td></tr>
  <tr><td><a href="<%="http://webflow/form/fe/index.asp?menu=report"%>"><img src="../images/orange_bullet.gif" border=0>Application for working in fatcory list report
</a></td></tr>

  <tr><td><a href="<%="http://webflow/form/fe/index.asp?menu=doc"%>"><img src="../images/orange_bullet.gif" border=0>Form Upload for working in fatcory list
</a></td></tr>


</table>
</body> </html> 

