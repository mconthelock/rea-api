<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="inc/system.inc"-->
<!--#include file="inc/db.inc"-->
<!--#include file="inc/connUtil.inc"-->
<!--#include file="inc/flowUtil.inc"-->
<!--#include file="inc/formUtil.inc"-->
<!--#include file="inc/date.inc"-->
<!--#include file="inc/authen.inc"-->
<!--#include file="inc/security.inc"-->
<!--#include file="inc/worklist.inc"-->

<%
Dim rsOra, sqlOra
Dim count, sign, formName
Dim empNo

if request.queryString("empNo") = "" then
	empNo = session("uid")
else
	empNo = request.queryString("empNo")
end if

set rsOra = createObject("adodb.recordset")

if session("uid") = "" then
	response.redirect "loginFrm.asp"
end if
%>

<html>
<head>
<link href="css/common.css" rel="stylesheet" type="text/css">
<style>
	a {font-family: ms sans serif; font-size: 10pt}
</style>
</head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">

<%
'Read form name
sqlOra = "select vName from formMst where nNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	formName = rsOra("vName")
else
%>
<table>
	<tr>
		<td>Form not found!</td>
	</tr>
</table>
<%
end if

rsOra.close
%>

  <table cellspacing="0" cellpadding="0" border="0" width="100%" align="center">
    <tr>
		<td bgcolor="#0000FF">
            
        <table border="0" cellspacing="1" cellpadding="2" width="100%">
          <tr bgcolor="#99CCEE"> 
            <td colspan="3"><font color="#000000"><b><%=formName%></b></font></td>
          </tr>
          <%if getAuth(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), session("uid"), AUTH_FORM_ADMIN) or getAuth(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), session("uid"), AUTH_KEY_OPR) then
		sqlOra = "select " _
		& "to_char(nNo) as nNo2, vOrgNo, cYear, vFormPage " _
		& "from " _
		& "formMst " _
		& "where " _
		& "nNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "'"
		'response.write sqlOra & "<br>"
		rsOra.open sqlOra, connOra
		
		if not rsOra.EOF then
			if inStr(rsOra("vFormPage"), "?") > 0 then
				sign = "&"
			else
				sign = "?"
			end if
	%>
          <tr>
            <td bgcolor="#FFCC00" align="center"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"><a href="<%=rsOra("vFormPage") & sign%>no=<%=rsOra("nNo2")%>&orgNo=<%=rsOra("vOrgNo")%>&y=<%=rsOra("cYear")%>&empNo=<%=session("uid")%>">Create 
              request</a></td>
            <td align="center" bgcolor="#FFFFCC"><font color="#000000">&nbsp;</font></td>
          </tr>
          <%
		end if
		
		rsOra.close
	end if
	%>
          <!--########## COMING SOON ##########-->
          <%count = countComing()%>
          <tr valign="top">
            <td bgcolor="#FFCC00" align="center"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"> 
              <%if count = 0 then%>
              <font color="#000000">Coming soon</font> 
              <%else%>
              <a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/workflow/waitList.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&empNo=<%=empNo%>">Coming soon</a> 
              <%end if%>
              <%
		if count > 0 then 
			'response.write "<br>"
			'showList STEP_WAIT, session("uid")
		end if
		%>
            </td>
            <td align="center" bgcolor="#FFFFCC"><font color="#000000">[<%=count%>]</font></td>
          </tr>
          <!--########## END COMING SOON ##########-->
          <!--########## READY ##########-->
          <%count = countReady()%>
          <tr valign="top">
            <td bgcolor="#FFCC00" align="center"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"> 
              <%if count = 0 then%>
              <font color="#000000">Wait for your approval</font> 
              <%else%>
              <a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/workflow/readyList.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&empNo=<%=empNo%>">Wait for your approval</a> 
              <%end if%>
              <%
		if count > 0 then 
			response.write "<br>"
			showList STEP_READY, session("uid")
		end if
		%>
            </td>
            <td align="center" bgcolor="#FFFFCC"><font color="#000000">[<%=count%>]</font></td>
          </tr>
          <!--########## END READY ##########-->
          <!--########## APPROVE ##########-->
          <%count = countApprove()%>
          <tr valign="top">
            <td bgcolor="#FFCC00" align="center"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"> 
              <%if count = 0 then%>
              <font color="#000000">Approve</font> 
              <%else%>
              <a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/workflow/approveList.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&empNo=<%=empNo%>">Approve</a> 
              <%end if%>
            </td>
            <td align="center" bgcolor="#FFFFCC"><font color="#000000">[<%=count%>]</font></td>
          </tr>
          <!--########## END APPROVE ##########-->
          <!--########## REPRESENTATIVE ##########-->
		  <%count = countRepApprove(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), session("uid"))%>
          <tr valign="top">
            <td bgcolor="#FFCC00" align="center"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"> 
			  <%count = countRepApprove(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), empNo)%>
              <%if count = 0 then%>
              <font color="#000000">Representative approve</font> 
              <%else%>
              <a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/workflow/repAList.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&empNo=<%=empNo%>">Representative approve</a> 
              <%end if%>
            </td>
            <td align="center" bgcolor="#FFFFCC"><font color="#000000">[<%=count%>]</font></td>
          </tr>
          <!--########## END REPRESENTATIVE ##########-->
          <!--########## MIND ##########-->
          <%
	count = countMine(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), session("uid"))
	%>
          <%if count > 0 then%>
          <tr valign="top">
            <td bgcolor="#FFCC00" align="center" width="1%"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/workflow/mineList.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&empNo=<%=empNo%>">Mine</a><%'if count > 0 then showMindList session("uid") end if%></td>
              <%else%>
            <td bgcolor="#FFCC00" align="center" width="1%"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"><font color="#000000">Mine</font></td>
            <%end if%>
            <td align="center" bgcolor="#FFFFCC"><font color="#000000">[<%=count%>]</font></td>
          </tr>
          <!--########## END MIND ##########-->
          <!--########## FINISH ##########-->
          <%
			count = countFinish(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), session("uid"))
			%>
          <tr valign="top">
            <td bgcolor="#FFCC00" align="center"><img src="imgs/doc.gif" width="20" height="20"></td>
            <td bgcolor="#FFCC00"> 
              <%if count = 0 then%>
              <font color="#000000">Finished</font> 
              <%else%>
              <a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/workflow/finList.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&empNo=<%=empNo%>">Finished</a> 
              <%end if%>
            </td>
            <td align="center" bgcolor="#FFFFCC"><font color="#000000">[<%=count%>]</font></td>
          </tr>
          <!--########## END FINISH ##########-->
        </table>
		</td>
	</tr>
</table>
<center>
    <a href="javascript: window.location.href = 'http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/index.asp'">Back to Forms</a>
  </center>
	</body> 
	</html> 
