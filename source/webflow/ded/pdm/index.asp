<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!--#include file="../../inc/hisUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/misc.inc"-->
<!--include file="inc/formUtil.inc"-->
<!--#include file="inc/local_pdm.inc"-->
<%
Dim requesterNo, requester, inputerNo, inputer
Dim mode, flowStatus
Dim rsOra, sqlOra
Dim no, orgNo, y, y2, runNo, apv, remark, requestTo, uid, pwd
Dim i
Dim out
Dim appNo
Dim jobController 

set rsOra = createObject("adodb.recordset")

'Get approver
if session("uid") = request.queryString("empNo") then
	apv = request.queryString("empNo")
else
	apv = session("uid")
end if

'Get mode
if apv <> "" then
	mode = getMode(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("y2"), request.queryString("runNo"), apv)
	
	if request.queryString("m") <> "" then
		mode = request.queryString("m")
	end if
	
	if mode = MODE_ADD and not (getAuth(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), apv, AUTH_KEY_OPR) or getAuth(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), apv, AUTH_FORM_ADMIN)) then
		mode = MODE_VIEW
	end if
else
	mode = MODE_VIEW
end if

'Get flow status
flowStatus = getFlowStatus(connOra, request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("y2"), request.queryString("runNo"))

if mode = MODE_ADD then
	inputerNo = session("uid")
	inputer = session("uName")
else
	sqlOra = "select vInputer, a.sName as sName, vReqNo, b.sName as sName2 from form, amec.aEmployee a, amec.aEmployee b " _
	& "where " _
	& "form.nFrmNo = '" & request.queryString("no") & "' and form.vOrgNo = '" & request.queryString("orgNo") & "' " _
	& "and form.cYear = '" & request.queryString("y") & "' and form.cYear2 = '" & request.queryString("y2") & "' " _
	& "and form.nRunNo = '" & request.queryString("runNo") & "' " _
	& "and a.sEmpNo = vInputer and b.sEmpNo = vReqNo " _
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, connOra
	
	if rsOra.EOF then
		response.write "Form data not found!"
		response.end
	else
		inputerNo = rsOra("vInputer")
		inputer = rsOra("sName")
		requesterNo = rsOra("vReqNo")
		requester = rsOra("sName2")
	end if
	
	rsOra.close
end if

if mode = MODE_ADD or mode = MODE_EDIT then
	no = request.queryString("no")
	orgNo = request.queryString("orgNo")
	y = request.queryString("y")
	y2 = request.queryString("y2")
	runNo = request.queryString("runNo")
	'apv = request.queryString("empNo")
else
	set rsOra = createObject("adodb.recordset")
	sqlOra = "select to_char(nFrmNo) as nFrmNo2, to_char(nRunNo) as nRunNo2, form.* from form where nFrmNo = '" & request.queryString("no") & "'" _
	 & " and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cYear2 = '" & request.queryString("y2") & "' and nRunNo = '" & request.queryString("runNo") & "'"
	rsOra.open sqlOra, connOra
	
	if not rsOra.EOF then
		no = rsOra("nFrmNo2")
		orgNo = rsOra("vOrgNo")
		y = rsOra("cYear")
		y2 = rsOra("cYear2")
		runNo = rsOra("nRunNo2")
'		apv = rsOra("apv")
	else
		response.write "Form data not found!"
		response.end
	end if
	
	rsOra.close
end if
jobController = isJobController(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("y2"), request.queryString("runNo"), session("uid"))
readFlowRepAuth no, orgNo, y, y2, runNo, session("uid")
%>
<html>
<head> 
<link href="../css/ep.css" rel="stylesheet" type="text/css">
<link href="../css/epTH.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../js/window.js"></script>
<script language="JavaScript" src="../../js/form.js"></script>
<script language="JavaScript">
function validate(act)
{
	return true;
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000"  leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0"> 
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/date.inc"-->
<!--#include file="../../inc/authen.inc"-->
<form name="frmMain" method="post" action="cgi-bin/action.asp" onSubmit="javascript: return <%if mode = MODE_ADD then%>validate()<%else%>confirm('Do you want to ' + document.frmMain.act.value + ' this form?')<%end if%>">
  <input type="hidden" name="apv" value="<%=apv%>">
  <input type="hidden" name="no" value="<%=no%>">
  <input type="hidden" name="orgNo" value="<%=orgNo%>">
  <input type="hidden" name="y" value="<%=y%>">
  <input type="hidden" name="y2" value="<%=y2%>">
  <input type="hidden" name="runNo" value="<%=runNo%>">
  <input type="hidden" name="act" value="">
  <input type="hidden" name="bp" value="<%=request.queryString("bp")%>">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
      <td align="center" bgcolor="#99CCFF"><b><font color="#000000">Local-PDM Registeration</font></b></td>
      <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#FFFFFF" abackground="../imgs/border/book/bg.gif"> 
        <table cellspacing="2" cellpadding="1" border="0" width="1">
          <%if mode <> MODE_ADD then%>
          <tr> 
            <td width="1%" nowrap height="25"><b><font color="#000000">Form No:</font></b></td>
            <td width="1%" height="25" nowrap><font color="#000000"><%=toFormNumber(no, orgNo, y, y2, runNo)%></font></td>
          </tr>
          <%end if%>
          <tr> 
            <td width="1%" nowrap height="25"><b><font color="#000000">Input by:</font></b></td>
            <td width="1%" height="25" nowrap><font color="#000000"><%="(" & inputerNo & ") " & inputer%></font></td>
          </tr>
          <tr> 
            <td width="1%" nowrap><b><font color="#000000">Requested by:</font></b></td>
            <td width="1%" nowrap><font color="#000000"> 
              <%
						if MODE = MODE_ADD then
						%>
              <input type="text" name="txtReq" title="Enter employee ID of requester">&nbsp;<font color="#FF0000">(Employee number)</font>
              <%
						else
							response.write "(" & requesterNo & ") " & requester
						end if
						%>
              </font></td>
          </tr>
        </table>
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr> 
            <td width="1%" colspan="2" nowrap valign="top"><img src="../../imgs/spacer.gif" height="4"></td>
          </tr>
          <tr> 
            <td width="1%" colspan="2" nowrap valign="top" bgcolor="#99CCFF"><b><font color="#000000"> 
              <%if mode = MODE_ADD then response.write "Please select " else response.write "Requested " end if%>
              Local-PDM Registration :</font></b></td>
          </tr>
          <%if mode = MODE_ADD then%>
          <!--########## ADD MODE DATA ##########-->
          <tr> 
            <td bgcolor="#0000FF"> 
              <table border="0" cellspacing="1" cellpadding="2" width="100%">
                <tr> 
                  <td align="center" nowrap colspan="3" width="0" bgcolor="#99CCFF"><b><font color="#000000">Request for:</font></b></td>
                  <td align="center" nowrap rowspan="2" bgcolor="#99CCFF"><b><font color="#000000">Authority Level</font></b></td>
                  <td align="center" nowrap width="0" rowspan="2" colspan="3" bgcolor="#99CCFF"><b><font color="#000000">Purpose</font></b> 
                  </td>
                </tr>
                <tr> 
                  <td align="center" nowrap width="0" bgcolor="#99CCFF"><b><font color="#000000">Add</font></b></td>
                  <td align="center" nowrap width="0" bgcolor="#99CCFF"><b><font color="#000000">Edit</font></b></td>
                  <td align="center" nowrap width="0" bgcolor="#99CCFF"><b><font color="#000000">Delete</font></b></td>
                </tr>
                <tr bgcolor="#FFFFCC">
   				  <td align="center" width="0"><font color="#000000"> 
                    <input type="radio" name="optReq1" value="01,1" checked>
                    </font></td>
                  <td align="center" width="0"><font color="#000000"> 
                    <input type="radio" name="optReq1" value="02,2">
                    </font></td>
                  <td align="center" width="0"><font color="#000000"> 
                    <input type="radio" name="optReq1" value="03,3">
                    </font></td>
                  <td nowrap width="0" align="left"> 
				  <%
								sqlOra = "select * from pdmauth order by vname" 
								rsOra.open sqlOra, connOra
					%>			
                    <select name="Authority">
                      <%
											do while not rsOra.EOF
					  %>
                      <option value="<%=rsOra("cNo")%>"><%=rsOra("vName")%></option>
                      <%
												rsOra.moveNext
											loop
											%>
                    </select>
                  </td>
                  <td nowrap colspan="3" align="center" width="0"> <font color="#000000"> 
                    <input type="text" name="Purpose" maxlength="4000">
                    </font> <font color="#000000"> </font> <font color="#000000"> 
                    </font></td>
                </tr>
                 <!--<tr> 
                   <td colspan="8" bgcolor="#FFFFFF" align="center"><br>
                     <font color="#000000"><b>Data not found!</b></font><br>
                     &nbsp;</td>
                 </tr>-->
                <%		
					rsOra.close
				%>
              </table>
            </td>
          </tr>
          <!--########## END ADD MODE DATA ##########-->
          <%elseif mode = MODE_EDIT and flowStatus = FLOW_RUNNING then%>
          <!--########## EDIT MODE DATA ##########-->
          <tr> 
            <td width="1%" colspan="2" nowrap valign="top" bgcolor="#0000FF"> 
              <table border="0" cellspacing="1" cellpadding="2" width="100%">
              <%
							sqlOra = "select vName, cReq, vPurpose, vUserno from pdmform a, pdmauth b " _
								& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
								& "and cYear = '" & request.queryString("y") & "' and cYear2 = '" & request.queryString("y2") & "' " _
								& "and nRunNo = '" & request.queryString("runNo") & "' and a.cNo = b.cNo" 
								 
									
								rsOra.open sqlOra, connOra
						
								if not rsOra.EOF then
								%>
								<tr bgcolor="#FFFFCC"> 
								  <td nowrap width="10%" bgcolor="#FFCC00"><font color="#000000">Request for:</font></td>
								  <td align="left" width="90%" nowrap><font color="#000000"> 
								<%
									if rsOra("cReq") = "1" then
										response.write "Add"
									elseif rsOra("cReq") = "2" then
										response.write "Edit"
									elseif rsOra("cReq") = "3" then
										response.write "Delete"
									else
										response.write "Unknown!"
									end if
									%>
                    </font></td>
                </tr>
				<tr bgcolor="#FFFFCC"> 
                  <td nowrap width="10%" bgcolor="#FFCC00"><font color="#000000">Authorized level:</font></td>
                  <td align="left" width="90%" nowrap><font color="#000000"> 
                    <%=rsOra("vName")%></font></td>
                </tr>
				<tr bgcolor="#FFFFCC"> 
                  <td nowrap width="10%" bgcolor="#FFCC00"><font color="#000000">Purpose:</font></td>
                  <td align="left" width="90%" nowrap><font color="#000000"> 
                    <%=rsOra("vPurpose")%></font></td>
                </tr>
                <% 
				else
				%>				
				
                <tr> 
                  <td colspan="9" bgcolor="#FFFFFF" align="center"><br>
                    <font color="#000000"><b>Data not found!</b></font><br>
                    &nbsp;</td>
                </tr>
                <%
								end if
								
								rsOra.close
								%>
              </table>
				<!--########## REMARK ##########-->
				<table width="100%" cellspacing="2" cellpadding="1" bgcolor="#FFFFFF">
				  <tbody><tr>
					<td>&nbsp;</td>
				  </tr>
				  <tr> 
					<td width="10%" nowrap="" valign="top"><b><font color="#000000">Remark:</font></b></td>
					<td valign="left" width="70%" nowrap=""> 
					  <textarea name="txtRemark" cols="50" rows="5" onkeyup="javascript: chkLng(document.frmMain.txtRemark, 1024, 'Remark is no longer than 1024 charactors')"></textarea>
					</td>
				  </tr>
				</tbody></table>
				<!--########## END REMARK ##########-->
            </td>
          </tr>
          <!--########## END EDIT MODE DATA ##########-->
          <%else%>
          <!--########## VIEW MODE DATA ##########-->
          <tr> 
            <td bgcolor="#0000FF"> 
              <table border="0" cellspacing="1" cellpadding="2" width="100%">
                 <%
								sqlOra = "select a.cNo, cReq, cApprove, b.vName as appName, c.vName as extData " _
								& "from " _
								& "darwinData a, darwinApp b, darwinExtData c " _
								& "where " _
								& "a.nFrmNo = '" & request.queryString("no") & "' and a.vOrgNo = '" & request.queryString("orgNo") & "' " _
								& "and a.cYear = '" & request.queryString("y") & "' and a.cYear2 = '" & request.queryString("y2") & "' " _
								& "and a.nRunNo = '" & request.queryString("runNo") & "' " _
								 & "and c.cNo(+) = a.cNo and c.cRunNo(+) = a.cRunNo and b.cNo = a.cNo " _
								 & "order by appName"

								sqlOra = "select vName, cReq, vPurpose from pdmform a, pdmauth b " _
								& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
								& "and cYear = '" & request.queryString("y") & "' and cYear2 = '" & request.queryString("y2") & "' " _
								& "and nRunNo = '" & request.queryString("runNo") & "' and a.cNo = b.cNo" 
								'response.write sqlOra & "<br>"
								rsOra.open sqlOra, connOra
								
								if not rsOra.EOF then
									
								%>
                <tr bgcolor="#FFFFCC"> 
                  <td nowrap width="10%" bgcolor="#FFCC00"><font color="#000000">Request for:</font></td>
                  <td align="left" width="90%" nowrap><font color="#000000"> 
                    <%
									if rsOra("cReq") = "1" then
										response.write "Add"
									elseif rsOra("cReq") = "2" then
										response.write "Edit"
									elseif rsOra("cReq") = "3" then
										response.write "Delete"
									else
										response.write "Unknown!"
									end if
									%>
                    </font></td>
                </tr>
				<tr bgcolor="#FFFFCC"> 
                  <td nowrap width="10%" bgcolor="#FFCC00"><font color="#000000">Authorized level:</font></td>
                  <td align="left" width="90%" nowrap><font color="#000000"> 
                    <%=rsOra("vName")%></font></td>
                </tr>
				<tr bgcolor="#FFFFCC"> 
                  <td nowrap width="10%" bgcolor="#FFCC00"><font color="#000000">Purpose:</font></td>
                  <td align="left" width="90%" nowrap><font color="#000000"> 
                    <%=rsOra("vPurpose")%></font></td>
                </tr>
                <%
								else
								%>
                <tr> 
                  <td colspan="5" bgcolor="#FFFFFF" align="center"><br>
                    <font color="#000000"><b>Data not found!</b></font><br>
                    &nbsp;</td>
                </tr>
                <%
								end if
								
								rsOra.close
								%>
              </table>
            </td>
          </tr>
          <!--########## END VIEW MODE DATA ##########-->
          <%end if%>
        </table>
        <input type="hidden" name="totalApp" value="<%=i%>">
        <!--########## REMARK ##########-->
        <%if (mode = MODE_ADD or mode = MODE_EDIT) and 1 = 2 then%>
        <table width="1" cellspacing="2" cellpadding="1">
          <tr> 
            <td width="1%" colspan="2" nowrap valign="top"><img src="../../imgs/spacer.gif" height="4"></td>
          </tr>
          <tr> 
            <td width="1%" nowrap valign="top"><b><font color="#000000">Remark:</font></b></td>
            <td valign="top" width="1%" nowrap> 
              <textarea name="txtRemark" cols="50" rows="5" onKeyUp="javascript: chkLng(document.frmMain.txtRemark, 1024, 'Remark is no longer than 1024 charactors')"></textarea>
            </td>
          </tr>
        </table>
        <%end if%>
        <!--########## END REMARK ##########-->

      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/bg.gif">&nbsp;</td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif"></td>
      <td align="left" background="../imgs/border/book/bg.gif"> 
        <%
			select case mode
				case MODE_ADD
			%>
        <center><input type="button" value="Request" onClick="javascript: if(validate('request')) {document.frmMain.act.value = 'request'; document.frmMain.submit()}">
        <%
				case MODE_EDIT
			%>
			<% if (jobController = true) then %>
					<table>
					<tr>
					<td><font color="#000000"><b>User :</b></font></td>
					<td><input type="text" value="" name = "user"></td>
					</tr>
					<tr> 
				    <td><font color="#000000"><b>Password :</b></font></td>
					<td><input type="text" value="" name = "pass"></td>
					</tr>
					</table>
					<td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
					</tr>
					<tr> 
					  <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
					  <td background="../imgs/border/book/bg.gif">&nbsp;</td>
					  <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
					</tr>
					<tr> 
					<td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif"></td>
					<td align="center" background="../imgs/border/book/bg.gif"> 
					<%
		end if %>
        <center><input type="button" value="Approve" onClick="javascript: if(validate('approve')) {document.frmMain.act.value = 'approve'; document.frmMain.submit()}">
		<input type="button" value="Reject" onClick="javascript: if(validate('reject')) {document.frmMain.act.value = 'reject'; document.frmMain.submit()}">
		<input type="button" value="Back" onClick="javascript: window.location.href = '<%
		if request.queryString("bp") = "" then
			response.write PAGE_DEFAULT_BACK
		else
			response.write request.queryString("bp")
		end if
		%>?no=<%=no%>&orgNo=<%=orgNo%>&y=<%=y%>'" title="Click here to back to previous page">
        <!--<!input type="submit" name="act" value="Reject all" title="Click here to reject this request">-->
        <%
			end select
			%>
        <%if mode = MODE_ADD then%>
        <input type="reset" value="Reset">
		<input type="button" value="Back" onClick="javascript: window.location.href = '<%
		if request.queryString("bp") = "" then
			response.write PAGE_DEFAULT_BACK
		else
			response.write request.queryString("bp")
		end if
		%>?no=<%=no%>&orgNo=<%=orgNo%>&y=<%=y%>'" title="Click here to back to previous page">
        <%end if%>
        <%
			if getAuth(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), session("uid"), AUTH_FORM_ADMIN) and flowStatus = FLOW_APPROVE and request.queryString("m") = "" then
			%>
        <center><input type="button" value="Save change" onClick="javascript: document.frmMain.act.value = 'save change'; document.frmMain.submit()" title="Click here to save change you made to this form">
        <input type="button" value="Back" onClick="javascript: window.location.href = '<%
		if request.queryString("bp") = "" then
			response.write PAGE_DEFAULT_BACK
		else
			response.write request.queryString("bp")
		end if
		%>?no=<%=no%>&orgNo=<%=orgNo%>&y=<%=y%>'" title="Click here to back to previous page">
		<%
			end if
			%>
		<%if mode = MODE_VIEW then%>
			<center><input type="button" value="Back" onClick="javascript: window.location.href = '<%
			if request.queryString("bp") = "" then
				response.write PAGE_DEFAULT_BACK
			else
				response.write request.queryString("bp")
			end if
			%>?no=<%=no%>&orgNo=<%=orgNo%>&y=<%=y%>'" title="Click here to back to previous page">
        <%
			end if
			%>
      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
    </tr>
    <tr> 
      <td align="right" valign="top"><img src="../imgs/border/book/blc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/brc.gif" width="17" height="18"></td>
    </tr>
  </table>
  <!--########## APPLICATION DESCRIPTION ##########-->
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <!--########## END APPLICATION DESCRIPTION ##########-->
  <%if mode = MODE_EDIT or mode = MODE_VIEW then%>
  <!--######### WORK FLOW ##########-->
  <br>
  <br>
  <%showFlow no, orgNo, y, y2, runNo, requester, inputer, "remark,legend"%>
  <!--######### END WORK FLOW ##########-->
  <%end if%>
</form>
<!--#include file="../../inc/copyRight.inc"-->
<%if mode = MODE_ADD then%>
<script language="JavaScript">
document.frmMain.txtReq.focus();
</script>
<br>
&nbsp; 
<%end if%>
</body>
</html>


