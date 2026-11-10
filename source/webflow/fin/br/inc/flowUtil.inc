<!--#include virtual="/form/inc/flowUtil2.inc"-->
<%
'##########
'# Function: createBrFlow(connOra, frmNo, orgNo, y, requester, inputer, remark, apvno , receiver)
'# Description: This function create flow (transaction) for specific form.
'# Argument: 
'# 1. connOra = Oracle connection.
'# 2. frmNo = form number.
'# 3. orgNo = organization number.
'# 4. y = 2 digit from right of year.
'# 5. requester = emp no. of requester .
'# 6. inputer = emp no. of inputer.
'# 7. remark = remark from requester or inputer (who is input form data).
'# 8. apvno = emp no. owner.
'# 9. receiver = emp no. of receiver.
'##########
function createBrFlow(connOra, frmNo, orgNo, y, requester, inputer, remark, apvno , receiver)
	Dim rsOra, rsOra2, rsOra3, sqlOra
	Dim repNo, mgr, mgrPos, mgrRep, runNo, reqOrgNo, stepSt, empNo
	Dim y2, createDate, createTime
	Dim flowStyle ' 1 = Have organize, 11 = Have organize and there are normal step
	
	set rsOra = createObject("adodb.recordset")
	set rsOra2 = createObject("adodb.recordset")
	set rsOra3 = createObject("adodb.recordset")

	'connOra.beginTrans
	
	'########## VALIDATE DATA ##########
	'Check existing requester
	if chkUser(requester) then
		response.write "Requester not found!"
		response.end
	end if

	'Check existing inputer
	if chkUser(inputer) then
		response.write "Inputer not found!"
		response.end
	end if
	
	'Check existing owner
	if chkUser(apvno) then
		response.write "Owner not found!"
		response.end
	end if
	
	'if receiver <> "N" then
		'Check existing recieve
		if chkUser(receiver) then
			response.write "Receiver not found!"
			response.end
		end if
	'end if
	'Check flow
	sqlOra = "select to_char(nFrmNo) as nFrmNo, flowMst.* from flowMst start with nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' " _
	& "and cStart = '1' connect by nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' " _
	& "and cStepNo = prior cStepNextNo"
'	response.write sqlOra & "<br>"
	rsOra.open sqlOra, connOra
	
	if rsOra.EOF then
		response.write "Flow master not found!"
		response.end
	end if
	
	rsOra.close
	'########## END VALIDATE DATA ##########

	'########## PREPARING DATA ##########
	remark = oraFormatStr(left(remark, 1024))
	y2 = year(date)
	createTime = toFullTime(time)
	createDate = toShortDate(date, "DD/MM/YYYY")
	'########## END PREPARING DATA ##########

	'########## CREATE FORM FOR INPUTER ##########
	'if 1 = 2 then
	'Get max runing number
	sqlOra = "select max(nRunNo) from form where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "'"
'	sqlOra = "select max(nRunNo) from form where nFrmNo = '" & frmNo & "' and cYear = '" & y & "' "
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, connOra

	if not rsOra.EOF then
		if not isNull(rsOra("max(nRunNo)")) then
			runNo = rsOra("max(nRunNo)") + 1
		end if
	end if
	
	if isNull(runNo) or runNo = "" then
		runNo = 1
	end if
	
	rsOra.close
	
	'Create form
	sqlOra = "insert into form (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, vReqNo, vInputer, vRemark, dReqDate, cReqTime, vFormPage) values ('" & frmNo & "', '" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & requester & "', '" & inputer & "', '" & remark & "', to_date('" & createDate & "', 'DD/MM/YYYY'), '" & createTime & "', (select vFormPage from formMst where nNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "'))"
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, connOra
	'########## END CREATE FORM FOR INPUTER ##########

	'########## CREATE FLOW ##########
	'Read flow master
	sqlOra = "select to_char(nFrmNo) as nFrmNo, flowMst.* from flowMst start with nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' " _
	& "and cStart = '1' connect by nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' " _
	& "and cStepNo = prior cStepNextNo"
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, connOra

	'Skip your position
	if not rsOra.EOF then
'		rsOra.moveNext
	end if
	
	'Create flow
	flowStyle = ""
	
	do while not rsOra.EOF
		if rsOra("cExtData") = "01"  then
			repNo = getRep(frmNo, orgNo, y, apvno)
			if isNull(repNo) then
				'repNo = rsOra2("sempno")
				repNo = apvno
			end if

			sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, cType, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
			& "values " _
			& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & apvno & "', '" & repNo & "', '" & APV_NONE & "', '" & rsOra("cType") & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
			'response.write "1 = "&sqlOra & "<br>"
			'response.end
			rsOra3.open sqlOra, connOra
		elseif rsOra("cExtData") = "02" then
				'if receiver = "N" then
				'	sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, cType, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
				'	& "values " _
				'	& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_NOT_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & rsOra("VAPVNO") & "', ' ', '" & APV_NONE & "', '" & rsOra("cType") & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
					
					'response.write sqlOra & "<br>"
					'response.end
				'	rsOra3.open sqlOra, connOra
				'else
					repNo = getRep(frmNo, orgNo, y, receiver)
					if isNull(repNo) then
						'repNo = rsOra2("sempno")
						repNo = receiver
					end if
					'response.write "rrrrrrrrrrrr = "&receiver&"ccccccccccc"
					sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, cType, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
					& "values " _
					& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & receiver & "', '" & repNo & "', '" & APV_NONE & "', '" & rsOra("cType") & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
					'response.write "JJJJJJJJJ="&sqlOra & "<br>"
					'response.end
					rsOra3.open sqlOra, connOra
				'end if
				'response.write "2 = "&sqlOra & "<br>"
		elseif rsOra("cExtData") = "03" then
					sqlOra = "select headno from sequenceOrg where empno = '" & receiver & "' and cco = '0'"
					rsOra3.open sqlOra, connOra		
					apvno = rsOra3("headno")
					repNo = getRep(frmNo, orgNo, y, apvno)
					if isNull(repNo) then
						repNo = apvno
					end if
					rsOra3.Close
					sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, cType, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
							& "values " _
							& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & apvno & "', '" & repNo & "', '" & APV_NONE & "', '" & rsOra("cType") & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
							'response.write "1 = "&sqlOra & "<br>"
							'response.end
							rsOra3.open sqlOra, connOra
		else
			'########## SET APPROVER ##########
			select case rsOra("cType")
					'########## REFER TO INPUTER (ORGANIZE) ##########
				case "1"
					'Set flow style flag
					if flowStyle = "" then
						flowStyle = "1"
					end if
					
					'Read inputer organization
					'reqOrgNo = getOrgNo(inputer)
					reqOrgNo = getOrgNo(requester)
					
					'Read approver in an organization
					sqlOra = "select * " _
					& "from orgPos " _
					& "where vOrgNo in (select distinct vOrgNo from orgTree start with vOrgNo = '" & reqOrgNo & "' connect by vOrgNo = prior vParent) and vPosNo = '" & rsOra("vPosNo") & "' " _
					& "and vEmpNo in (select headNo from sequenceOrg start with empNo = '" & requester & "' and sPosCode = '" & getPos(requester) & "' and cCo = '0' " _
					& "connect by prior headNo = empNo and prior sPosCode1 = sPosCode)"

					'sqlOra = "select * from orgPos where vOrgNo in (select distinct vOrgNo from orgTree start with vOrgNo = '" & reqOrgNo & "' connect by vOrgNo = prior vParent) and (vPosNo = '" & rsOra("vPosNo") & "' and vPosNo < (select sPosCode from amec.aEmployee where sEmpNo = '" & requester & "')) "
					'response.write sqlOra & "<br>"
					
					rsOra2.open sqlOra, connOra
					
					if not rsOra2.EOF then
						'Set flow style
						flowStyle = "11"

						do while not rsOra2.EOF
							'Read approver's representative refer to inputer
							repNo = getRep(frmNo, orgNo, y, rsOra2("vEmpNo"))
							
							if isNull(repNo) then
								repNo = rsOra2("vEmpNo")
							end if
		
							'Add step to transaction
							sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, cType, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
							& "values " _
							& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & rsOra2("vEmpNo") & "', '" & repNo & "', '" & APV_NONE & "', '" & rsOra("cType") & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
							'response.write sqlOra & "<br>"
							rsOra3.open sqlOra, connOra
							
							rsOra2.moveNext
						loop
		
						rsOra2.close
					else
						rsOra2.close
	
						sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
						& "values " _
						& "('" & frmNo & "','" & rsOra("vOrgNo") & "', '" & rsOra("cYear") & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_NOT_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & requester & "', '" & requester & "', '" & APV_NONE & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
						'response.write sqlOra & "<br>"
						rsOra2.open sqlOra, connOra
					end if
					'########## END REFER TO INPUTER (ORGANIZE) ##########

					'########## REFER TO FORM OWNER ##########
				case "2"
					'Read current approver in an organization refer to position and organization
					sqlOra = "select vEmpNo from orgPos where vPosNo = '" & rsOra("vPosNo") & "' and vOrgNo = '" & rsOra("vApvOrgNo") & "' order by vEmpNo"
					'response.write sqlOra & "<br>"
					rsOra2.open sqlOra, connOra
			
					if not rsOra2.EOF then
						do while not rsOra2.EOF
							repNo = getRep(frmNo, orgNo, y, rsOra2("vEmpNo"))
							
							if isNull(repNo) then
								repNo = rsOra2("vEmpNo")
							end if
							
							'Add step to transaction
							sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, cType, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
							& "values " _
							& "('" & frmNo & "','" & rsOra("vOrgNo") & "', '" & rsOra("cYear") & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & rsOra2("vEmpNo") & "', '" & repNo & "', '" & APV_NONE & "', '" & rsOra("cType") & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
							'response.write sqlOra & "<br>"
							rsOra3.open sqlOra, connOra
							
							rsOra2.moveNext
						loop
						
						rsOra2.close
					else
						rsOra2.close
						
	'					sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO) " _
	'					& "values " _
	'					& "('" & frmNo & "','" & rsOra("vOrgNo") & "', '" & rsOra("cYear") & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_NOT_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & inputer & "', '" & inputer & "', '" & APV_NONE & "')"

						sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, vPosNo, VAPVNO, VREPNO, CAPVSTNO, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
						& "values " _
						& "('" & frmNo & "','" & rsOra("vOrgNo") & "', '" & rsOra("cYear") & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & STEP_NOT_USE & "', '" & rsOra("cStart") & "', '" & rsOra("vPosNo") & "', '" & requester & "', '" & requester & "', '" & APV_NONE & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
						'response.write sqlOra & "<br>"
						rsOra3.open sqlOra, connOra
					end if
					'########## END REFER TO FORM OWNER ##########

					'########## REFER TO EMPLOYEE NUMBER ##########
				case "3"
					'if not isNull(rsOra("vApvNo")) then
					if rsOra("vApvNo") <> "'" & DEFAULT_APPROVER & "'" or (rsOra("vApvNo") = "'" & DEFAULT_APPROVER & "'" and not isNull(rsOra("cExtData"))) then
						empNo = rsOra("vApvNo")
						repNo = getRep(frmNo, orgNo, y, empNo)
						
						if isNull(repNo) then
							repNo = empNo
						end if

						stepSt = STEP_USE
					else
						'empNo = inputer
						'repNo = inputer

						empNo = requester
						repNo = requester
						stepSt = STEP_NOT_USE
					end if

					sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, VAPVNO, VREPNO, CAPVSTNO, cType, vUrl, cExtData, cApvType, cRejType, cApplyAll) " _
					& "values " _
					& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & rsOra("cStepNo") & "', '" & rsOra("cStepNextNo") & "', '" & stepSt & "', '" & rsOra("cStart") & "', '" & oraFormatStr(empNo) & "', '" & repNo & "', '" & APV_NONE & "', '" & rsOra("cType") & "', '" & oraFormatStr(left(rsOra("vUrl"), 1024)) & "', '" & oraFormatStr(left(rsOra("cExtData"), 2)) & "', '" & rsOra("cApvType") & "', '" & rsOra("cRejType") & "', '" & rsOra("cApplyAll") & "')"
					'response.write sqlOra & "<br>"
					rsOra2.open sqlOra, connOra
					'########## END REFER TO EMPLOYEE NUMBER ##########
			end select
			'########## END SET APPROVER ##########
		end if
		
		rsOra.moveNext
	loop

	rsOra.close
	'########## END CREATE FLOW ##########
	
	
	
	'#if receiver <> "N" then
	'#	sqlOra = "update flow set cStepSt = '"&STEP_NOT_USE&"' where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cExtData='06'"
	'#	rsOra.open sqlOra, connOra
		'response.write "sqlOra not use ="&sqlOra
	'##end if
	'########## UPDATE STATUS ##########
	'Update form status
	sqlOra = "update form set cSt = '" & FLOW_ON_GOING & "' where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "'"
	rsOra.open sqlOra, connOra
	
	'Delete isn't use step
	sqlOra = "select to_char(nFrmNo) as nFrmNo2, to_char(nRunNo) as nRunNo2, flow.* from flow where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStepSt = '" & STEP_NOT_USE & "'"
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, connOra
	
	do while not rsOra.EOF
		deleteFlowStep connOra, rsOra("nFrmNo2"), rsOra("vOrgNo"), rsOra("cYear"), rsOra("cYear2"), rsOra("nRunNo2"), "flow", rsOra("cStepNo"), rsOra("cStepNextNo")
		rsOra.moveNext
	loop
	
	rsOra.close
	
	'########## ADD MANAGER ##########
	'Add inputer's manager to flow if this flow has organize step but no match manager.
	if flowStyle = "1" then
		'Read manager
		mgr = getMgr(requester)

		'Do only if requester has manager and requester is not directory and division manager
		'response.write isNull(mgr) & "-" & isMgr(requester)
		'response.end
		if not isNull(mgr) and not isMgr(requester) then
			
			'Read manager's respresentative
			mgrRep = getRep(frmNo, orgNo, y, getMgr(requester))
			
			if isNull(mgrRep) then
				mgrRep = mgr
			end if
			
			sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, VAPVNO, VREPNO, CAPVSTNO, cType, dApvDate, cApvTime, vRemark, vUrl, cApplyAll) " _
			& "values " _
			& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & STEP_POS_SPECIAL & "', (select distinct cStepNo from flow where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStart = '1'), '" & STEP_NORMAL & "', '1', '" & mgr & "', '" & mgrRep & "', '" & APV_NONE & "', '0', null, null, null, (select vFormPage from formMst where nNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "'), '" & APPLY_ALL_BOTH & "')"
			'response.write sqlOra & "<br>"
			rsOra2.open sqlOra, connOra

			'Change current start node flag
			sqlOra = "update flow set cStart = '0' " _
			& "where " _
			& "nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' " _
			& "and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' " _
			& "and cStart = '1' and cStepSt = '" & STEP_NORMAL & "' " _
			& "and cStepNo <> '" & STEP_POS_SPECIAL & "'"
			'response.write sqlOra & "<br>"
			rsOra2.open sqlOra, connOra	
		end if
	end if
	'########## END ADD MANAGER ##########

	'########## ADD FIRST STEP ##########
	'Dermpan (4-DEC-2002): อันที่จริงคิดว่าควรจะทำขั้นตอนนี้ก่อนแล้วกำหนดให้ทุก step มีค่า cStart = '0' จะง่ายกว่าและลดขั้นตอนการเปลี่ยน cStart ด้านล่าง
	'Add inputer step to head of flow
	sqlOra = "insert into flow (NFRMNO, VORGNO, CYEAR, cYear2, nRunNo, CSTEPNO, CSTEPNEXTNO, cStepSt, cStart, VAPVNO, VREPNO, CAPVSTNO, cType, dApvDate, cApvTime, vRemark, vUrl, cApplyAll) " _
	& "values " _
	& "('" & frmNo & "','" & orgNo & "', '" & y & "', '" & y2 & "', " & runNo & ", '" & STEP_POS_FIRST & "', (select distinct cStepNo from flow where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStart = '1'), '" & STEP_READY & "', '1', '" & requester & "', '" & requester & "', '" & APV_NONE & "', '0', to_date('" & createDate & "', 'DD/MM/YYYY'), '" & createTime & "', '" & remark & "', (select vFormPage from formMst where nNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "'), '" & APPLY_ALL_BOTH & "')"
	'response.write sqlOra & "<br>"
	rsOra2.open sqlOra, connOra

	'Change current start node flag
	sqlOra = "update flow set cStart = '0' " _
	& "where " _
	& "nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' " _
	& "and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' " _
	& "and cStart = '1' and cStepSt = '" & STEP_NORMAL & "'"
	'response.write sqlOra & "<br>"
	rsOra2.open sqlOra, connOra	
	'########## END ADD FIRST STEP ##########

	'Update flow status for next approver
	sqlOra = "update flow set cApvStNo = '" & APV_NONE & "', cStepSt = '" & STEP_READY & "' where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStepNo = (select cStepNextNo from flow where nFrmNo = '" & frmNo & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' and cYear2 = '" & y2 & "' and nRunNo = '" & runNo & "' and cStart = '1') and cStepSt = '" & STEP_USE & "'"
	'rsOra.open sqlOra, connOra
	
	'เปลี่ยนจากการกำหนดสถานะตอนสร้างโฟล์มาเป็นการ approve แทนเพื่อให้การทำงานเป็นไปตามครรลองคลองธรรมนะ
	approve connOra, frmNo, orgNo, y, y2, runNo, requester, remark
	'########## END UPDATE STATUS ##########
	
	'connOra.commitTrans
'response.end
	set rsOra =nothing
	set rsOra2 =nothing
	set rsOra3 =nothing
	
	createBrFlow = y2 & "," & runNo
end function


%>
