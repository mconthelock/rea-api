<%
Option Explicit
%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->
<!--#include file="../inc/report.inc"-->
<!--#include file="../inc/reportData.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->

<%
Dim rsOra, rsOraPos, sqlOra, sqlOraPos
Dim i, orgType
Dim pageSize, pageCount, absolutePage, recordCount, currentRecord  'Page/Report control variable
Dim whereClause, orderClause
Dim pos, posNo, posLst
Dim secCode, depCode, divCode

set rsOra = createObject("adodb.recordset")
set rsOraPos = createObject("adodb.recordset")

'Set page size
if request.queryString("ps") = ""  then
	pageSize = 10
else
	pageSize = cint(request.queryString("ps"))
end if

if pageSize < 0 then
	pageSize = 10
end if

'Set absolute page
if request.queryString("ab") = "" then
	absolutePage = 1
else
	absolutePage = cint(request.queryString("ab"))
end if

if absolutePage < 0 then
	absolutePage = 1
end if

'########## GENERATE WHERE CLAUSE ##########
'whereClause = " sEmpNo in (select distinct vEmpNo from orgPos) and cStatus <> '0'"
whereClause = "cStatus <> '0'"

if request.queryString("search") = "1" then
	if request.queryString("txtUid") <> "" then
		if inStr(request.queryString("txtUid"), "*") > 0 or inStr(request.queryString("txtUid"), "?") > 0 then
			whereClause = whereClause & " and sEmpNo like '" & replace(replace(request.queryString("txtUid"), "*", "%"), "?", "_") & "'"
		else
			whereClause = whereClause & " and sEmpNo like '%" & request.queryString("txtUid") & "%'"
		end if
	end if
	
	if request.queryString("txtUName") <> "" then
		if inStr(request.queryString("txtUName"), "*") > 0 or inStr(request.queryString("txtUName"), "?") > 0 then
			whereClause = whereClause & " and sName like '" & replace(replace(request.queryString("txtUName"), "*", "%"), "?", "_") & "'"
		else
			whereClause = whereClause & " and sName like '%" & request.queryString("txtUName") & "%'"
		end if
	end if
	
	if request.queryString("selSec") <> "" then
		whereClause = whereClause & " and sSecCode = '" & request.queryString("selSec") & "'"
	end if
	
	if request.queryString("selDep") <> "" then
		whereClause = whereClause & " and sDepCode = '" & request.queryString("selDep") & "'"
	end if
	
	if request.queryString("selDiv") <> "" then
		whereClause = whereClause & " and sDivCode = '" & request.queryString("selDiv") & "'"
	end if
else
	whereClause = "sSecCode = '" & request.queryString("orgNo") & "' or sDepCode = '" & request.queryString("orgNo") & "' or sDivCode = '" & request.queryString("orgNo") & "'"
end if

whereClause = whereClause & " and cStatus <> '0'"
orderClause = "order by sEmpNo"
'########## END GENERATE WHERE CLAUSE ##########

'########## READ POSITION ##########
sqlOraPos = "select * from amec.pPosition order by sPosCode"
rsOraPos.open sqlOraPos, connOra

if not rsOraPos.EOF then
	posLst = rsOraPos("sPosCode") & "&&" & rsOraPos("sPosition")

	rsOraPos.moveNext
	do while not rsOraPos.EOF
		posLst = posLst & "||" & rsOraPos("sPosCode") & "&&" & rsOraPos("sPosition")
		rsOraPos.moveNext
	loop
else
	response.write "Position not found!"
end if

rsOraPos.close
'########## END READ POSITION ##########
%>
<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
<script language="JavaScript" src="../js/window.js"></script>
<title>Add organization</title>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%showPageLoading%>
<form name="frmMain" method="get" action="cgi-bin/addOrgEmp.asp">
	<input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
    
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
            <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            
      <td bgcolor="#FFFFFF"> 
        <table width="100%" border="0" cellspacing="2" cellpadding="1">
                    <tr> 
                        <td colspan="2" nowrap bgcolor="#AACCEE" align="center"><b><font color="#000000">Add 
                            employee to organization<br>
                            "<%=request.queryString("orgNo")%>: <%=request.queryString("orgName")%> (<%=request.queryString("orgDesc")%>)"</font></b></td>
                    </tr>
                    <tr valign="top"> 
                        <td width="1%" nowrap><b><font color="#000000">Employee:</font></b></td>
						<td></td>
					</tr>
					
					<!--########## EMPLOYEE LIST ##########-->
					<tr>
                        <td colspan="2" align="center">
							<table width="100%" border="0" cellspacing="0" cellpadding="0">
								<tr>
									<td bgcolor="#0000FF" align="center">
										<table width="100%" border="0" cellspacing="1" cellpadding="2">
                                            <tr class="hdrRow" align="center"> 
                                                <td width="1%" nowrap><b><font color="#000000">&nbsp;</font></b></td>
                                                <td width="1%" nowrap><b><font color="#000000">No</font></b></td>
                                                <td width="1%" nowrap><b><font color="#000000">Emp 
                                                    No</font></b></td>
                                                <td nowrap width="100%"><b><font color="#000000">Name</font></b></td>
                                                <td nowrap width="1%"><b><font color="#000000">Sec</font></b></td>
                                                <td nowrap width="1%"><b><font color="#000000">Dep</font></b></td>
                                                <td nowrap width="1%"><b><font color="#000000">Div</font></b></td>
                                                <td nowrap width="1%"><b><font color="#000000">Position</font></b></td>
                                            </tr>
                                            <%
										sqlOra = "select sEmpNo, sName, sPosCode, sSecCode, sDepCode, sDivCode from amec.aEmployee where " & whereClause & " " & orderClause
										'response.write sqlOra & "<br>"
										rsOra.open sqlOra, connOra
										
										if not rsOra.EOF then
											i = 1
											
											'Goto start record
											recordCount = 1
											do while not rsOra.EOF and recordCount  < absolutePage * pageSize - pageSize + 1
												recordCount = recordCount + 1
												rsOra.moveNext
											loop
										
											currentRecord = absolutePage * pageSize - pageSize + 1
										
											i = currentRecord
											
											do while not rsOra.EOF and i < absolutePage * pageSize + 1
												'response.write "<input type=" & QUOTE & "checkbox" & QUOTE & " name=" & QUOTE & "chkChild" & QUOTE & " value=" & QUOTE & rsOra("vOrgNo") & QUOTE & ">" & rsOra("vOrgNo") & ": " & rsOra("vName") & " (" & rsOra("vDesc") & ")</option><br>"
										%>
                                            <tr bgcolor="<%
												if i mod 2 = 0 then
													response.write "#FFCC00"
												else
													response.write "#FFFFCC"
												end if
												%>"> 
                                                <td align="center" width="1%" nowrap> 
                                                    <input type="checkbox" name="chkEmp" value="<%=rsOra("sEmpNo")%>">
                                                </td>
                                                <td align="center" width="1%"><font color="#000000"><%=i%></font></td>
                                                <td align="center" width="1%" nowrap><font color="#000000"><%=rsOra("sEmpNo")%></font></td>
                                                <td nowrap width="100%"><font color="#000000"><%=rsOra("sName")%></font></td>
                                                <td nowrap width="1%" align="center"><font color="#000000"><%
												if rsOra("sSecCode") = secCode then
													response.write "&nbsp;"
												else
													secCode = rsOra("sSecCode")
													response.write secCode
												end if%></font></td>
                                                <td nowrap width="1%" align="center"><font color="#000000"><%
												if rsOra("sDepCode") = depCode then
													response.write "&nbsp;"
												else
													depCode = rsOra("sDepCode")
													response.write depCode
												end if%></font></td>
                                                <td nowrap width="1%" align="center"><font color="#000000"><%
												if rsOra("sDivCode") = divCode then
													response.write "&nbsp;"
												else
													divCode = rsOra("sDivCode")
													response.write divCode
												end if%></font></td>
                                                <td nowrap width="1%"><%
												if posLst <> "" then
												%>
                                                    <select name="selPos">
                                                        <%
													for each pos in split(posLst, "||")
														posNo = left(pos, inStr(posLst, "&&") - 1)
												%>
                                                        <option value="<%=rsOra("sEmpNo") & "&&" & posNo%>" <%if rsOra("sPosCode") = posNo then response.write "selected" end if%>><%=trim(mid(pos, inStr(pos, "&&") + 2, len(pos)))%></option>
                                                        <%
													next
												%>
                                                    </select>
                                                    <%
												else
													response.write "Position not found!"
												end if
												%></td>
                                            </tr>
                                            <%
												i = i + 1
												rsOra.moveNext
											loop
										else
										%>
                                            <tr> 
                                                
                        <td bgcolor="#FFFFFF" align="center" nowrap width="1%" colspan="9"><b><font color="#FF0000"><br>
                          Employee not found!</font></b><br>
                          &nbsp; </td>
                                            </tr>
                                            <%
										end if
										
										rsOra.close
										%>
                                        </table>
								</td>
							</tr>
							</table>
                        </td>
                    </tr>
					<!--########## END EMPLOYEE LIST ##########-->

                </table>
            </td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
        </tr>

		<!--########## PAGE NAVIGATOR ##########-->
		<tr> 
			<td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
			<td abackground="../imgs/ctlbg.jpg" align="center"> 
				<%
				'##### Display page navigator #####
				sqlOra = "select count(*) as recordCount from amec.aEmployee where " & whereClause
				'response.write sqlOra & "<br>"
				rsOra.open sqlOra, connOra, adOpenStatic, adLockReadOnly, adCmdText
				
				'		recordCount = currentRecord - 1 
				'recordCount = rsOra("recordCount")
				'Place Here
				displayPageNavigator absolutePage, rsOra("recordCount"), pageSize, request.serverVariables("SCRIPT_NAME") & "?orgNo=" & request.queryString("orgNo") & "&hDisplayData=1" & "&txtUid=" & request.queryString("txtUid") & "&txtUName=" & request.queryString("txtUName") & "&selSec=" & request.queryString("selSec") & "&selDep=" & request.queryString("selDep") & "&selDiv=" & request.queryString("selDiv") & "&search=" & request.queryString("search") & "&orgName=" & server.urlEncode(request.queryString("orgName") & "") & "&orgDesc=" & server.urlEncode(request.queryString("orgDesc") & "") & "&ab="
				rsOra.close
				%>
			</td>
			<td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
		</tr>
		<!--########## END PAGE NAVIGATOR ##########-->

        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/bg.gif" align="center">&nbsp;</td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif"></td>
            <td align="center" background="../imgs/border/book/bg.gif"> 
                <input type="submit" value="Add">
                <input type="button" value="Back" onClick="javascript: window.location.href='editOrgFrm.asp?orgNo=<%=request.queryString("orgNo")%>'">
            </td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="../imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
</form>

<!--######### SEARCH FORM ##########-->
<table align="center">
<form name="frmSearchUsr" action="<%=request.serverVariables("SCRIPT_NAME")%>" mothod="get">
	<input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
	<input type="hidden" name="orgName" value="<%=request.queryString("orgName")%>">
	<input type="hidden" name="orgDesc" value="<%=request.queryString("orgDesc")%>">
	<input type="hidden" name="search" value="1">
	<tr> 
		<td colspan="3"  abackground="../imgs/ctlbg.jpg"> 
			<hr size="1">
		</td>
	</tr>
	<tr abackground="../imgs/ctlbg.jpg"> 
		<td colspan="3"> 
			<table border="0" cellspacing="0" cellpadding="0">
				<tr> 
					<td colspan="2"><font class="FontText">Employee 
						search</font></td>
				</tr>
				<tr> 
					<td width="1%" nowrap><font class="FontText">Employee ID:</font></td>
					<td> 
						<input type="text" name="txtUid" class="inputNormal" value="<%=request.queryString("txtUid")%>">
					</td>
				</tr>
				<tr> 
					<td width="1%" nowrap><font class="FontText">Name</font></td>
					<td> 
						<input type="text" name="txtUName" class="inputNormal" value="<%=request.queryString("txtUName")%>">
					</td>
				</tr>
				<tr> 
					<td width="1%" nowrap><font class="FontText">Section:</font></td>
					<td> 
						<%
						sqlOra = "select sSecCode, sSection from amec.pSection order by sSection"
						rsOra.open sqlOra, connOra
						
						if not rsOra.EOF then
						%>
						<select name="selSec" class="inputNormal">
							<option value=""></option>
							<%do while not rsOra.EOF%>
							<option value="<%=rsOra("sSecCode")%>" <%if request.queryString("selSec") = rsOra("sSecCode") then response.write "selected" end if%>><%=rsOra("sSection")%></option>
							<%
								rsOra.moveNext
							loop
							%>
						</select>
						<%
		else
			response.write "Section data not found!"
		end if
		
		rsOra.close
		%>
					</td>
				</tr>
				<tr> 
					<td width="1%" nowrap><font class="FontText">Department:</font></td>
					<td> 
						<%
		sqlOra = "select sDepCode, sDepartment from amec.pDepartment order by sDepartment"
		rsOra.open sqlOra, connOra
		
		if not rsOra.EOF then
		%>
						<select name="selDep" class="inputNormal">
							<option value=""></option>
							<%do while not rsOra.EOF%>
							<option value="<%=rsOra("sDepCode")%>" <%if request.queryString("selDep") = rsOra("sDepCode") then response.write "selected" end if%>><%=rsOra("sDepartment")%></option>
							<%
				rsOra.moveNext
			loop
			%>
						</select>
						<%
		else
			response.write "Department data not found!"
		end if
		
		rsOra.close
		%>
					</td>
				</tr>
				<tr> 
					<td width="1%" nowrap><font class="FontText">Division:</font></td>
					<td> 
						<%
		sqlOra = "select sDivCode, sDivision from amec.pDivision order by sDivision"
		rsOra.open sqlOra, connOra
		
		if not rsOra.EOF then
		%>
						<select name="selDiv" class="inputNormal">
							<option value=""></option>
							<%do while not rsOra.EOF%>
							<option value="<%=rsOra("sDivCode")%>" <%if request.queryString("selDiv") = rsOra("sDivCode") then response.write "selected" end if%>><%=rsOra("sDivision")%></option>
							<%
				rsOra.moveNext
			loop
			%>
						</select>
						<%
			else
				response.write "Division data not found!"
			end if
			
			rsOra.close
			%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr> 
		<td colspan="3"  abackground="../imgs/ctlbg.jpg" align="right"><img src="/form/imgs/spacer.gif" height="4"></td>
	</tr>
	<tr> 
		<td colspan="3"  abackground="../imgs/ctlbg.jpg" align="right"> 
			<input class="inputNormal" type="button" value="List all" onClick="javascript: window.location.href = '<%=request.serverVariables("SCRIPT_NAME")%>?orgNo=<%=request.queryString("orgNo")%>'">
			<input class="inputNormal" type="submit" value="Search">
		</td>
	</tr>
</form>
</table>
<!--######### END SEARCH FORM ##########-->
<!--#include file="../inc/copyRight.inc"-->
<%hidePageLoading%>
</body>
</html>
