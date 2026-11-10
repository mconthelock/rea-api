<!--#include file="inc/system.inc"-->
<!--#include file="inc/db.inc"-->
<!--#include file="inc/connUtil.inc"-->
<html>
<head>
<meta http-equiv="refresh" content="60;url=index.asp">
</head>
<body>
<%
Dim i, rsOra, rsOra2, sqlOra
Dim sign, url

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")
%>

<pre>
<b>Screen 2</b><hr size="1">
<b>Login user: </b><%="(" & session("uid") & ") " & session("uName")%> <a href="logout.asp">logout</a>
<form action="login.asp">
Uid: <input type="text" name="uid" value="<%=session("uid")%>">
Pwd: <input type="text" name="pwd" value="<%=request.queryString("pwd")%>"><input type="submit" name="Submit" value="Submit">
</form>
<hr size="1">
<%if 1 = 2 then%>
<b>S c r e e n</b>
<a href="ep/darwin/darwinRegFrm.asp">DARWIN user registration</a>
<a href="ep/eMailReg/eMailRegFrm.asp">E-Mail user registration</a>
<a href="ep/wwwReg/wwwRegFrm.asp">WWW user registration</a>
<a href="ep/psr/psrRegFrm.asp">Parts supply system user registration</a>
<hr noshade size="1">
<%end if%>
<b>M a s t e r</b>
<a href="datmgr/stepLst.asp">Step master</a>
<a href="datmgr/orgLst.asp">Organization master</a>
<a href="datmgr/frmLst.asp">Form master</a>
<!--<a href="datmgr/posLst.asp">Position master</a>-->
<hr noshade size="1">
<b>F o r m </b><a href="index.asp">Screen 1 (Test)</a>
<%
sqlOra = "select formMst.*, to_char(nNo) as nNo from formMst order by vName"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	response.write "<ol>"

	do while not rsOra.EOF
			if inStr(rsOra("vFormPage"), "?") = 0 then
				sign = "?"
			else
				sign = "&"
			end if
			
		'Display form name
		if not isNull(rsOra("vFormPage")) and session("uid") <> "" then
			response.write "<li>" & rsOra("vName") & "<ul type=square><i><a href=" & rsOra("vFormPage") & sign & "no=" & rsOra("nNo") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear") & "&empNo=" & session("uid") & ">Create request</a></i></ul>"
		else
			response.write "<li>" & rsOra("vName") & "<br>"
		end if

		'########## DISPLAY MIND TRANSATION ##########
		sqlOra = "select distinct form.nRunNo, to_char(flow.nRunNo) as nRunNo2, to_char(flow.nFrmNo) as nFrmNo2, flow.vOrgNo, flow.cYear, flow.cYear2, flow.cStepSt, flow.vApvNo, form.vInputer, formMst.vFormPage, vUrl " _
		& "from flow, form, formMst, flowMst " _
		& "where " _
		& "form.nFrmNo = '" & rsOra("nNo") & "' and form.vOrgNo = '" & rsOra("vOrgNo") & "' and form.cYear = '" & rsOra("cYear") & "' and flow.nFrmNo = form.nFrmNo and flow.vOrgNo = form.vOrgNo and flow.cYear = form.cYear " _
		& "and form.vReqNo = '" & session("uid") & "' " _
		& "and flow.cYear2 = form.cYear2 and flow.nRunNo = form.nRunNo " _
		& "and formMst.nNo = form.nFrmNo and formMst.vOrgNo = form.vOrgNo and formMst.cYear = form.cYear " _
		& "and flowMst.nFrmNo = form.nFrmNo and flowMst.vOrgNo = form.vOrgNo and flowMst.cYear = form.cYear " _
		& "and cSt = '" & FLOW_RUNNING & "' " _
		& "order by nRunNo desc"
		'response.write sqlOra & "<br>"
		rsOra2.open sqlOra, connOra
		
		if not rsOra2.EOF then
			i = 1
			response.write "<ul type=square><i>Mind</i>"
			do while not rsOra2.EOF
				if rsOra2("vUrl") <> rsOra2("vFormPage") and not isNull(rsOra2("vUrl")) then
					url = rsOra2("vUrl")
				else
					url = rsOra2("vFormPage")
				end if
				
				if rsOra2("cStepSt") = STEP_WAIT or rsOra2("cStepSt") = STEP_READY and session("uid") <> "" then
			%><li><a href="<%=url & sign & "no=" & rsOra2("nFrmNo2") & "&orgNo=" & rsOra2("vOrgNo") & "&y=" & rsOra2("cYear") & "&y2=" & rsOra2("cYear2") & "&runNo=" & rsOra2("nRunNo2") & "&empNo=" & rsOra2("vApvNo")%>">Transaction <%=i & ". Apv=" & rsOra2("vApvNo") & ", record=" & rsOra2("nRunNo2")%></a> <%=rsOra2("vName")%></li><%
				else
					response.write "<li>Transaction " & i & ". " & rsOra2("vApvNo") & " " & rsOra2("nRunNo2") & " " & rsOra2("vName") & "</li>"
				end if
				i = i + 1
				rsOra2.moveNext
			loop

			response.write "</ul>"
		end if
				
		rsOra2.close
		'########## END DISPLAY MIND TRANSATION ##########

		'########## DISPLAY COMING SOON TRANSATION ##########
		'########## END DISPLAY COMING SOON TRANSATION ##########

		if 1 = 1 then
		'Display running transaction
		sqlOra = "select distinct form.nRunNo, to_char(flow.nRunNo) as nRunNo2, to_char(flow.nFrmNo) as nFrmNo2, flow.vOrgNo, flow.cYear, flow.cYear2, flow.cStepSt, flow.vApvNo, form.vInputer, formMst.vFormPage, vUrl, stepStatus.vName from flow, form, formMst, stepStatus, flowMst " _
		& "where " _
		& "form.nFrmNo = '" & rsOra("nNo") & "' and form.vOrgNo = '" & rsOra("vOrgNo") & "' and form.cYear = '" & rsOra("cYear") & "' and flow.nFrmNo = form.nFrmNo and flow.vOrgNo = form.vOrgNo and flow.cYear = form.cYear " _
		& " and flow.cYear2 = form.cYear2 and flow.nRunNo = form.nRunNo " _
		& "and formMst.nNo = form.nFrmNo and formMst.vOrgNo = form.vOrgNo and formMst.cYear = form.cYear " _
		& "and flowMst.nFrmNo = form.nFrmNo and flowMst.vOrgNo = form.vOrgNo and flowMst.cYear = form.cYear " _
		& "and cStepSt = stepStatus.cNo and cStepSt in ('" & STEP_READY & "', '" & STEP_WAIT & "') " _
		& "order by nRunNo desc"
		rsOra2.open sqlOra, connOra
		
		if not rsOra2.EOF then
			i = 1
			response.write "<ul type=square><i>Running</i>"
			do while not rsOra2.EOF
				if rsOra2("vUrl") <> rsOra2("vFormPage") and not isNull(rsOra2("vUrl")) then
					url = rsOra2("vUrl")
				else
					url = rsOra2("vFormPage")
				end if
				
				if rsOra2("cStepSt") = STEP_WAIT or rsOra2("cStepSt") = STEP_READY and session("uid") <> "" then
			%><li><a href="<%=url & sign & "no=" & rsOra2("nFrmNo2") & "&orgNo=" & rsOra2("vOrgNo") & "&y=" & rsOra2("cYear") & "&y2=" & rsOra2("cYear2") & "&runNo=" & rsOra2("nRunNo2") & "&empNo=" & rsOra2("vApvNo")%>">Transaction <%=i & ". Apv=" & rsOra2("vApvNo") & ", record=" & rsOra2("nRunNo2")%></a> <%=rsOra2("vName")%></li><%
				else
					response.write "<li>Transaction " & i & ". " & rsOra2("vApvNo") & " " & rsOra2("nRunNo2") & " " & rsOra2("vName") & "</li>"
				end if
				i = i + 1
				rsOra2.moveNext
			loop

			response.write "</ul>"
		end if
				
		rsOra2.close

		'Display finished transaction
		sqlOra = "select distinct to_char(flow.nRunNo) as nRunNo2, to_char(flow.nFrmNo) as nFrmNo2, flow.vOrgNo, flow.cYear, flow.cYear2, flow.cStepSt, flow.vApvNo, form.vInputer, formMst.vFormPage, vUrl, flowStatus.vName from flow, form, formMst, flowMst, flowStatus " _
		& "where " _
		& "form.nFrmNo = '" & rsOra("nNo") & "' and form.vOrgNo = '" & rsOra("vOrgNo") & "' and form.cYear = '" & rsOra("cYear") & "' and flow.nFrmNo = form.nFrmNo and flow.vOrgNo = form.vOrgNo and flow.cYear = form.cYear " _
		& " and flow.cYear2 = form.cYear2 and flow.nRunNo = form.nRunNo " _
		& "and formMst.nNo = form.nFrmNo and formMst.vOrgNo = form.vOrgNo and formMst.cYear = form.cYear " _
		& "and flowMst.nFrmNo = form.nFrmNo and flowMst.vOrgNo = form.vOrgNo and flowMst.cYear = form.cYear " _
		& "and cSt = flowStatus.cNo and flow.cStepNextNo = '" & STEP_POS_LAST & "' and cSt in ('" & FLOW_APPROVE & "', '" & FLOW_REJECT & "') " _
		& "order by nRunNo2 desc"
		'response.write sqlOra & "<br>"
		rsOra2.open sqlOra, connOra
		
		if not rsOra2.EOF then
			i = 1
			response.write "<ul type=square><i>Finished</i>"
			do while not rsOra2.EOF
				if rsOra2("vUrl") <> rsOra2("vFormPage") and not isNull(rsOra2("vUrl")) then
					url = rsOra2("vUrl")
				else
					url = rsOra2("vFormPage")
				end if
				
				'if rsOra2("cStepSt") = STEP_WAIT or rsOra2("cStepSt") = STEP_READY then
			%><li><a href="<%=url & sign & "no=" & rsOra2("nFrmNo2") & "&orgNo=" & rsOra2("vOrgNo") & "&y=" & rsOra2("cYear") & "&y2=" & rsOra2("cYear2") & "&runNo=" & rsOra2("nRunNo2") & "&empNo=" & rsOra2("vApvNo")%>">Transaction <%=i & ". Adm=" & rsOra2("vApvNo") & ", record=" & rsOra2("nRunNo2")%></a> <%=rsOra2("vName")%></li><%
				'else
				'	response.write "<li>Transaction " & rsOra2("vApvNo") & " " & i & " " & rsOra2("vName") & "</li>"
				'end if
				
				i = i + 1
				rsOra2.moveNext
			loop

			response.write "</ul>"
		end if
				
		rsOra2.close
		end if
		
		rsOra.moveNext
	loop
	
	response.write "</ol>"
else
	response.write "Form not found!"
end if

rsOra.close
%>
<hr size="1">
<a href="table.asp">Table list</a>
<script language="JavaScript">
document.write( document.lastModified)
</script>
</pre>
</body>
</html>