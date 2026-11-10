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
%>

<html>
<head>
<!meta http-equiv="refresh" content="60;url=index.asp">
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<link href="../css/common.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">

<table><tr><td></td>
    <td><b>Create Development & Design Electronic Forms</b></td>
  </tr></table><%
sqlOra = "select " _
& "formMst.*, to_char(nNo) as nNo2 " _
& "from formMst " _
& "where " _
& "cStatus = '" & FORM_STS_SHOW & "' and vOrgNo like '0701%'" _
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
			
		if 	((rsOra("nNo2")= "6") and (rsOra("vOrgNo")="070101") and (rsOra("cYear")="19")) then
			rsOra.moveNext
		end if
		'Display form name
		if not isNull(rsOra("vFormPage")) and session("uid") <> "" then
			if rsOra("vName") = "T-manual release form" then
		%>
		 <tr>
		  <td><img src="../images/orange_bullet.gif"></td>
		  <td>
		    <a href="<%=rsOra("vFormPage") & sign & "no=" & rsOra("nNo2") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear") & "&bp=" & protocol & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/" & PS_SUB_SYSTEM_FOLDER & server.urlEncode("/create.asp?orgNo=") & rsOra("vOrgNo")%>"><%=rsOra("vName")%></a>
	<%	if (not isNull(rsOra("vDesc")) or rsOra("vDesc") <> "") then
			response.write " <font color=""gray"">(" & rsOra("vDesc") & ")</font>"
		end if	%>
			
		  </td>
		 </tr>
		 <tr>
		  <td><img src="../images/orange_bullet.gif"></td>
		  <td>
				<a href="javascript: var winAtch = window.open('<%=protocol%><%=request.serverVariables("SERVER_NAME") %>/<%=SYSTEM_FOLDER%>/<%=DED_SUB_SYSTEM_FOLDER%>/tm/Templete_T-Manual.xlsm', 'winAtch'); winAtch.focus(); void(0)">T-manual data Template</a>
		  </td>
		 </tr>
		 <tr>
		  <td><img src="../images/orange_bullet.gif"></td>
		  <td>
				<a href="javascript: var winAtch = window.open('<%=protocol%><%=request.serverVariables("SERVER_NAME") %>/<%=SYSTEM_FOLDER%>/<%=DED_SUB_SYSTEM_FOLDER%>/tm/Templete_T-Manual_Escalator.xlsm', 'winAtch'); winAtch.focus(); void(0)">T-manual data Template for Escalator</a>
		  </td>
		 </tr>
	
		<%  else %>
		<tr>
		  <td><img src="../images/orange_bullet.gif"></td>
		  <td>
		    <a href="<%=rsOra("vFormPage") & sign & "no=" & rsOra("nNo2") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear") & "&bp=" & protocol & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/" & PS_SUB_SYSTEM_FOLDER & server.urlEncode("/create.asp?orgNo=") & rsOra("vOrgNo")%>"><%=rsOra("vName")%></a>
	<%	if (not isNull(rsOra("vDesc")) or rsOra("vDesc") <> "") then
			response.write " <font color=""gray"">(" & rsOra("vDesc") & ")</font>"
		end if	%>
			
		  </td>
		 </tr>
		<%
			end if
		else
			response.write "<li>" & rsOra("vName") '& "<br>"
		end if
		
		rsOra.moveNext
	loop
else
	response.write "<tr><td>Form not found!</td></tr>"
end if

rsOra.close
response.write "</table>"
%>

<!--
<hr size="1">
<b><u>Master</u><br>Computer Accessory</b>
<a href="ep/comAcs/pgMstFrm.asp">Product group master</a><br>
<b>File Server</b>
<a href="ep/fileSvr/svrMstFrm.asp">Server master</a>
<a href="ep/fileSvr/fileMstFrm.asp">File master</a>
-->
<%
function showCS2()
	Dim rsOra2, sqlOra
	Dim i, sign
	Dim protocol
	If LCase(Request.ServerVariables("HTTPS")) = "on" Then
		protocol = "https://"
	Else
		protocol = "http://"
	End If
	set rsOra2 = createObject("adodb.recordset")
	
	'Display finished transaction
	sqlOra = "select distinct to_char(flow.nRunNo) as nRunNo2, to_char(flow.nFrmNo) as nFrmNo2, flow.vOrgNo, flow.cYear, flow.cYear2, flow.cStepSt, flow.vApvNo, form.vInputer, formMst.vFormPage, flowMst.vUrl, flowStatus.vName, vOrgNo2 from flow, form, formMst, flowMst, flowStatus, comSupData " _
	& "where " _
	& "form.nFrmNo = '10' and form.vOrgNo = '050603' and form.cYear = '02' " _
	& "and comSupData.nFrmNo = '10' and comSupData.vOrgNo = '050603' and comSupData.cYear = '02' and comSupData.cYear2 = form.cYear2 and comSupData.nRunNo = form.nRunNo " _
	& "and flow.nFrmNo = form.nFrmNo and flow.vOrgNo = form.vOrgNo and flow.cYear = form.cYear " _
	& "and flow.cYear2 = form.cYear2 and flow.nRunNo = form.nRunNo " _
	& "and formMst.nNo = form.nFrmNo and formMst.vOrgNo = form.vOrgNo and formMst.cYear = form.cYear " _
	& "and flowMst.nFrmNo = form.nFrmNo and flowMst.vOrgNo = form.vOrgNo and flowMst.cYear = form.cYear " _
	& "and flowMst.cStepNo = flow.cStepNo and flowMst.cStepNextNo = flow.cStepNextNo " _
	& "and cSt = flowStatus.cNo and flow.cStepNextNo = '" & STEP_POS_LAST & "' and cSt = '" & FLOW_APPROVE & "' " _
	& "order by nRunNo2 desc"
	'response.write sqlOra & "<br>"
	rsOra2.open sqlOra, connOra
	
	if not rsOra2.EOF then
		i = 1
		response.write "<ul type=square>"
		do while not rsOra2.EOF
			if rsOra2("vUrl") <> rsOra2("vFormPage") and not isNull(rsOra2("vUrl")) then
				url = rsOra2("vUrl")
			else
				url = rsOra2("vFormPage")
			end if

			if inStr(url, "?") = 0 then
				sign = "?"
			else
				sign = "&"
			end if
		%>
<li><a href="<%=protocol%><%=request.serverVariables("SERVER_NAME")%>/form/ep/comSup/index2.asp?<%="no=" & rsOra2("nFrmNo2") & "&orgNo=" & rsOra2("vOrgNo") & "&y=" & rsOra2("cYear") & "&y2=" & rsOra2("cYear2") & "&runNo=" & rsOra2("nRunNo2") & "&empNo=" & rsOra2("vApvNo")%>">Transaction <%=i & ". Section=" & rsOra2("vOrgNo2") & ", record=" & rsOra2("nRunNo2")%></a></li><%
			i = i + 1
			rsOra2.moveNext
		loop
		response.write "</ul>"
	end if
			
	rsOra2.close
end function
%>

<br>
<table>
  <tr><td></td></tr>
  <tr><td><b>Report</b></td></tr>
  <tr><td><a href="<%=protocol%>webflow.mitsubishielevatorasia.co.th/form/ded/dsa/report.asp"><img src="../images/orange_bullet.gif" border=0>AMEC's Self Drawing Check list
</a></td></tr>
</table>

</body> 
</html> 

