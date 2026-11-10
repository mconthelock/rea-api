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
<%
  Dim protocol
  If LCase(Request.ServerVariables("HTTPS")) = "on" Then
    protocol = "https://"
  Else
    protocol = "http://"
  End If
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
& "formMst.*, to_char(nNo) as nNo2, substr(vformpage,instr(vformpage, 'sr=')+3,1) as menu " _
& "from formMst " _
& "where " _
& "cStatus = '" & FORM_STS_SHOW & "' and vOrgNo like '0510%' " _
& "order by vName"
'& "cStatus = '" & FORM_STS_SHOW & "' and vOrgNo like '0510%' and VANAME like 'FE%'  " _
'response.write sqlOra
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	response.write "<table cellpadding=1 cellspacing=0 border=0>"
	do while not rsOra.EOF
			if menu <> rsOra("menu") then 
				i = 1
			end if
			if inStr(rsOra("vFormPage"), "?") = 0 then
				sign = "?"
			else
				sign = "&"
			end if
			
		'Display form name
		if(rsOra("menu") = "o") then
		response.write "<tr><td><img src=""../images/orange_bullet.gif""></td><td>"

		
		if not isNull(rsOra("vFormPage")) and session("uid") <> "" and (getAuth(rsOra("nNo2"), rsOra("vOrgNo"), rsOra("cYear"), session("uid"), FORM_AUTH_ADMIN) or getAuth(rsOra("nNo2"), rsOra("vOrgNo"), rsOra("cYear"), session("uid"), FORM_KEY_OPR)) then
		%>

		<!-- 2. <a href="<%=rsOra("vFormPage") & sign & "no=" & rsOra("nNo2") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear")  & "&empno=" & session("uid") & "&bp=" & "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/" & MAT_SUB_SYSTEM_FOLDER & "/create.asp"%>"><%=rsOra("vName")%></a> -->
		
    <%
    ' 1. ตรวจสอบและแทนที่ protocol ของ vFormPage ให้เป็น protocol เดียวกับหน้าที่กำลังเปิดอยู่
    Dim formUrl
    formUrl = Trim(rsOra("vFormPage") & "")

    If LCase(Left(formUrl, 8)) = "https://" Then
        formUrl = protocol & Mid(formUrl, 9)
    ElseIf LCase(Left(formUrl, 7)) = "http://" Then
        formUrl = protocol & Mid(formUrl, 8)
    End If
    %>
    <!-- 2. เปลี่ยนตรง &bp= จากเดิมที่เป็น "http://" ให้ใช้ตัวแปร protocol แทน -->
    <a href="<%=formUrl & sign & "no=" & rsOra("nNo2") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear") & "&empno=" & session("uid") & "&bp=" & protocol & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/" & MAT_SUB_SYSTEM_FOLDER & "/create.asp"%>"><%=rsOra("vName")%></a>


		
		<%
		else
			response.write rsOra("vName")
		end if

		if not isNull(rsOra("vDesc")) or rsOra("vDesc") <> "" then
			response.write " <font color=""gray"">(" & rsOra("vDesc") & ")</font>"
		end if

		response.write "</td></tr>"
		end if
		rsOra.moveNext
		
		
	loop
else
	response.write "<tr><td>Form not found!</td></tr>"
end if
rsOra.close
%>	
</table>



<%
'Working in factory master Management
if (getAuth("2", "051001", "20", session("uid"), AUTH_FORM_ADMIN)) then 
%>
<table>
  <tr><td></td></tr>
  <tr><td><b>Working in factory master Management</b></td></tr>
  <tr><td><a href="<%="//webflow.mitsubishielevatorasia.co.th/form/fe/index.asp?menu=subcon"%>"><img src="../images/orange_bullet.gif" border=0>Sub-contractor master</a></td></tr>
  <tr><td><a href="<%="//webflow.mitsubishielevatorasia.co.th/form/fe/index.asp?menu=subcon_worker"%>"><img src="../images/orange_bullet.gif" border=0>Sub-contractor worker master</a></td></tr>
  <tr><td><a href="<%="//webflow.mitsubishielevatorasia.co.th/fe/index.asp?menu=gas"%>"><img src="../images/orange_bullet.gif" border=0>Gas master</a></td></tr>
  <tr><td><a href="<%="//webflow.mitsubishielevatorasia.co.th/form/fe/index.asp?menu=gas_sub"%>"><img src="../images/orange_bullet.gif" border=0>Gas supervisor</a></td></tr>
  <tr><td><a href="#"><img src="../images/orange_bullet.gif" border=0>Hot work scope master</a></td></tr>
  <tr><td><a href="<%="//webflow.mitsubishielevatorasia.co.th/form/fe/index.asp?menu=item_mstd"%>"><img src="../images/orange_bullet.gif" border=0>Checking item master</a></td></tr>
  <tr><td><a href=#"><img src="../images/orange_bullet.gif" border=0>Machine installation and Construction inspection sheet</a></td></tr>
</table>
<%end if%>
<%
'Waste Disposal master management

'if (getAuth("14", "051001", "20", session("uid"), AUTH_FORM_ADMIN)) then 
'If Session("uid") = "15199" Or Session("uid") = "07004" Or Session("uid") = "13246" Then

if (getAuth("2", "051001", "20", session("uid"), AUTH_FORM_ADMIN)) then 
%>
<table>
  <tr><td></td></tr>
  <tr><td><b>Waste Disposal master management</b></td></tr>
  <!--
  <tr><td><a href="<%="http://webflow.mitsubishielevatorasia.co.th/form/fe/index2.asp?menu=vendor"%>"><img src="../images/orange_bullet.gif" border=0>Vendor Master</a></td></tr>
  <tr><td><a href="<%="http://webflow.mitsubishielevatorasia.co.th/form/fe/index2.asp?menu=wastelist"%>"><img src="../images/orange_bullet.gif" border=0>Waste List master</a></td></tr>
  <tr><td><a href="<%="http://webflow.mitsubishielevatorasia.co.th/form/fe/index2.asp?menu=product"%>"><img src="../images/orange_bullet.gif" border=0>Product master</a></td></tr>
  -->
  <tr><td><a href="<%="//amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/WasteDisaposal/show_factory_manager/"%>"><img src="../images/orange_bullet.gif" border=0>Waste disposal company Management</a></td></tr>
  
</table>
<%end if%>

<table>
  <tr><td></td></tr>
  <tr><td><b>Report</b></td></tr>
  <tr><td><a href="<%="//webflow.mitsubishielevatorasia.co.th/form/fe/index.asp?menu=report"%>"><img src="../images/orange_bullet.gif" border=0>Application for working in fatcory list report
</a></td></tr>
  <tr><td><b>Document</b></td></tr>
  <tr><td><a href="<%="//webflow.mitsubishielevatorasia.co.th/form/fe/index.asp?menu=doc"%>"><img src="../images/orange_bullet.gif" border=0>Doucument for working in fatcory list
</a></td></tr>


</table>
</body> </html> 

