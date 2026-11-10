<%
'##########
'# Function: showFlowExcel(connOra, frmNo, orgNo, y, y2, runNo, requester, inputer)
'# Description: This function display flow data of specific form.
'# Argument: 
'# 1. connOra = Oracle connection.
'# 2. frmNo = form number.
'# 3. orgNo = organization number.
'# 4. y = 2 digit from right of year.
'# 5. y2 = year.
'# 6. runNo = transaction running number.
'# 7. requester = employee number, who is an requester.
'# 8. inputer = employee number, who is an form inputer.
'##########
function showFlowExcel(frmNo, orgNo, y, y2, runNo, requester, inputer, style)
	Dim rsOra, sqlOra
	Dim styleRemark, i
	Dim showStepStatus
	
	styleRemark = cBool(inStr(style, "remark") > 0)
	
	set rsOra = createObject("adodb.recordset")
	
	sqlOra = "select distinct lv, to_char(nFrmNo) as nFrmNo2, to_char(nRunNo) as nRunNo2, a.*, sName, sPosition, vName " _
	& "from " _
	& "(select level as lv, flow.* from flow start with nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStart = '1' connect by nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStepNo = prior cStepNextNo) a " _
	& ", amec.aEmployee, amec.pPosition, stepMst " _
	& "where " _
	& "sEmpNo = vApvNo and amec.pPosition.sPosCode = amec.aEmployee.sPosCode and cNo = cStepNo"

	sqlOra = "select distinct lv, to_char(a.nFrmNo) as nFrmNo2, to_char(a.nRunNo) as nRunNo2, a.*, sName, sPosition, vName, cSt, a.vRemark " _
	& " from " _
	& " form, (select level as lv, flow.* from flow start with nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStart = '1' connect by nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStepNo = prior cStepNextNo) a " _
	& " , amec.aEmployee, amec.pPosition, stepMst " _
	& " where " _
	& " sEmpNo = vApvNo and amec.pPosition.sPosCode = amec.aEmployee.sPosCode and cNo = cStepNo " _
	& " and form.nFrmNo = a.nFrmNo and form.vOrgNo = a.vOrgNo and form.cYear = a.cYear and form.cYear2 = a.cYear2 and form.nRunNo = a.nRunNo " _
	& " order by a.lv"

	
'	sqlOra = "select to_char(a.nFrmNo) as nFrmNo2, to_char(a.nRunNo) as nRunNo2, a.*, sName, sPosition, vName, cSt " _
'	& "from " _
'	& "form, (select flow.* from flow where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "') a " _
'	& ", amec.aEmployee, amec.pPosition, stepMst " _
'	& "where " _
'	& "sEmpNo = vApvNo and amec.pPosition.sPosCode = amec.aEmployee.sPosCode and cNo = cStepNo " _
'	& "and form.nFrmNo = a.nFrmNo and form.vOrgNo = a.vOrgNo and form.cYear = a.cYear and form.cYear2 = a.cYear2 and form.nRunNo = a.nRunNo " 
'	& " order by vposno desc"

'	response.write sqlOra & "<br>"
'	response.end
	rsOra.open sqlOra, connOra
'	response.end
	%>
<table width="100%" border="0">	
  <tr>
  	  <td>&nbsp;</td>
	  <td>	
		  <table border="1" cellspacing="1" cellpadding="2" align="center">
			<tr bgcolor="#AACCEE"> 
			  <td align="center" nowrap width="1%"><b><font color="#000000">Status</font></b></td>		
			  <td align="center" nowrap width="1%"><b><font color="#000000">Step</font></b></td>
			  <td align="center" width="1%" nowrap><b><font color="#000000">Emp-no</font></b></td>
			  <td align="center" width="1%" nowrap><b><font color="#000000">Emp-name</font></b></td>
			  <td align="center" width="1%" nowrap><b><font color="#000000">Date</font></b></td>
			  <td align="center" width="1%" nowrap><b><font color="#000000">Time</font></b></td>
			  <%if styleRemark then%>
			  <td align="center" width="1%" nowrap><b><font color="#000000">Remark</font></b></td>
			  <%end if%>
			</tr>
			<%
		if not rsOra.EOF then
			i = 0
			do while not rsOra.EOF
				'Set show approve status flag, this is fool to do.
				if (rsOra("cStepNextNo") <> STEP_POS_LAST) or (rsOra("cSt") = FLOW_RUNNING and (rsOra("cStepSt") = STEP_WAIT or (rsOra("cStepSt") = STEP_READY or rsOra("cStepSt") = STEP_RETURN  or rsOra("cStepSt") = STEP_APPROVE))) or (rsOra("cSt") = FLOW_APPROVE) or (rsOra("cSt") = FLOW_REJECT and rsOra("cStepNextNo") = STEP_POS_LAST and (rsOra("cStepSt") = STEP_REJECT or rsOra("cStepSt") = STEP_DIE)) then
					showStepStatus = true
				else
					showStepStatus = false
				end if
	
				if rsOra("cStepSt") <> STEP_NOT_USE then
	%>
			<tr bgcolor="#FFFFF"> 
			  <td width="1%" align="center" nowrap><font color="#000000">
				<%
				if showStepStatus then
					if rsOra("cStepSt") = STEP_READY or rsOra("cStepSt") = STEP_RETURN then
						response.write "Waiting for approval"
					elseif rsOra("cStepSt") = STEP_WAIT then
						response.write "Coming soon"				
					else
						select case rsOra("cApvStNo")
							case APV_APPROVE
								if rsOra("cStepSt") = STEP_SKIP then
									response.write "Approved by other"				
								elseif rsOra("cStepSt") = STEP_DIE then
									response.write "Step was disable"				
								else
									response.write "Approved"				
								end if
							case APV_REJECT
								if rsOra("cStepSt") = STEP_SKIP then
									response.write "Rejected by other"											
								elseif rsOra("cStepSt") = STEP_DIE then
									response.write "Step was disable"																		
								else
									response.write "Rejected"																		
								end if
							case APV_RETURN
								response.write "Returned to requester"																								
							case else
								response.write "&nbsp;"
								if debug then
									response.write rsOra("cApvStNo")
								end if
							end select
					end if
				else
					response.write "&nbsp;"
				end if
			%>
				</font></td>
			  <td width="1%" nowrap><font color="#000000"><%=rsOra("vName")%></font></td>
			  <td width="1%" align="center" nowrap class="txt"> 
				<%	if rsOra("vRepNo") = rsOra("vApvNo") then %>
				<%=rsOra("vApvNo")%>
				<%else%>
				<%if showStepStatus then%>
				<%if rsOra("vRealApv") = rsOra("vApvNo") then %>
				<img src="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/imgs/stepstatus/approver.gif" alt="Approved by <%=rsOra("vRealApv")%>"> 
				<%end if%>
				<%end if%>
				<%=rsOra("vApvNo")%>
				<%if showStepStatus then%>
				<%if rsOra("vRealApv") = rsOra("vRepNo") then %>
				<img src="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/imgs/stepstatus/approver.gif" alt="Approved by <%=rsOra("vRealApv")%>"> 
				<%end if%>
				<%end if%>
				<%=rsOra("vRepNo")%>
				<%end if%>
			  </td>
			  <td width="1%" nowrap><font color="#000000"><%=rsOra("sName")%></font></td>
			  <td width="1%" align="center" nowrap><font color="#000000"> 
				<%if showStepStatus then%>
				<%=showData(toMedDate(rsOra("dApvDate")))%> 
				<%
			  else
				response.write "&nbsp;"
				end if
				%>
				</font></td>
			  <td width="1%" align="center" nowrap><font color="#000000"> 
				<%if showStepStatus then%>
				<%=showData(rsOra("cApvTime"))%> 
				<%
			  else
				  response.write "&nbsp;"
			  end if
			  %>
				</font></td>
			  <td width="1%" align="center" nowrap><font color="#000000"><%=rsOra("vRemark")%></font></td>
			</tr>
			<%
					i = i + 1
				end if
				rsOra.moveNext
			loop
		end if
		%>
			<tr> 
			  <td align="center" bgcolor="#FFFFFF" colspan="<%if styleRemark then response.write "7" else response.write "6" end if%>"><%=getFlowStatusName(connOra, frmNo, orgNo, y, y2, runNo)%></td>
			</tr>
		  </table>
	  </td>		  
  </tr>	  
</table>  
<%
end function
%>