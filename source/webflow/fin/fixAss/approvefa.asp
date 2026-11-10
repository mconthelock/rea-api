<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/connUtil.inc"--> 
<!--#include file="../../inc/db.inc"--> 
<!--#include file="../../inc/authen.inc"--> 
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!--#include file="../../inc/date.inc"-->
<%
	Dim rsOra , sqlOra , chkempty , repNo , asid , remark , reason , statusapp
	Dim rsOra2 , rsCon , rsCon2 , rsCon3 , chk , chead
	Dim sec , dept , div
	chkempty = true
	'response.write "test = "&connOra
	
	
	set rsOra = createObject("adodb.recordset")
	set rsOra2 = createObject("adodb.recordset")
	set rsCon = createObject("adodb.recordset") 
	set rsCon2 = createObject("adodb.recordset") 
	set rsCon3 = createObject("adodb.recordset") 
	connOra.BeginTrans
	
	
	select case lCase(request("act"))
	case "reject"
		'response.write "connOra = "&connOra&"<br/>"
		'response.write "no = "&request("no")&"<br/>"
		'response.write "orgNo = "&request("orgNo")&"<br/>"
		'response.write "y = "&request("y")&"<br/>"
		'response.write "y2 = "&request("y2")&"<br/>"
		'response.write "runno = "&request("runNo")&"<br/>"
		'response.write "apv = "&session("uid")&"<br/>"
		'response.end
		  'reject connOra, request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid") , request("txtRemark")
		  'reject connOra, "reject", request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid"), request("txtRemark")
		 Call reject(connOra, request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid"), request("txtRemark"))
	case "approve" 
	'	sqlOra = "select cExtdata from flow where nfrmno = '" & request("no") & "' and vorgno = '" & request("orgNo") & "'" _
		'	& " and cyear = '" & request("y") & "' and cyear2='" & request("y2") & "' and nrunno = '" & request("runNo") & "'" _
		'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') order by cExtdata ASC"
		'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cstepst='" & STEP_READY & "'"
		'	response.write sqlOra
		'	response.end
		'	rsOra.open sqlOra, connOra
		'	response.write "cExtdata"&request("stepExt")
			
			if request("stepExt") = "01" then
				if chkWorkingUser(request("picno")) then
%>
						<script language="javascript">
							alert("PIC not found")
							history.back();
						</script>
<%		
					response.end
				end if 
				sqlOra = "select sseccode , sdepcode , sdivcode from amec.aemployee where sempno = '"&session("uid")&"'"
				rsCon.open sqlOra, connOra
				sec = rsCon("sseccode")
				dept = rsCon("sdepcode") 
				div = rsCon("sdivcode")
				rsCon.close
				chk = 0
				if sec <> "00" then
					sqlOra = "select count(sempno) rnum from  amec.aemployee where sempno = '"&request("picno")&"' and sseccode = '"&sec&"'"
					rsCon.open sqlOra, connOra
					if rsCon("rnum") = 1 then
						chk = 1
					end if
					rsCon.close
				else
					if dept <> "00" then
							sqlOra = "select count(sempno) rnum from  amec.aemployee where sempno = '"&request("picno")&"' and sdepcode = '"&dept&"'"
							rsCon.open sqlOra, connOra
							if rsCon("rnum") = 1 then
								chk = 1
							end if
							rsCon.close
					else
						if div <> "00" then
							sqlOra = "select count(sempno) rnum from  amec.aemployee where sempno = '"&request("picno")&"' and sdivcode = '"&div&"'"
							rsCon.open sqlOra, connOra
							if rsCon("rnum") = 1 then
								chk = 1
							end if
							rsCon.close
						end if
					end if
				
				end if
			chead = 0
			sqlOra = "select count(empno) rnum from sequenceorg where empno = '"&request("picno")&"' and headno = '"&session("uid")&"'"
			
			
			rsCon.open sqlOra, connOra
			if rsCon("rnum") > 0 then
						chead = 1
			end if
			rsCon.close
			
			if chk = 0 and chead = 0 then
%>
						<script language="javascript">
							alert("PIC not in subordination")
							history.back();
						</script>
<%		
					response.end
				end if 	
				
			end if
	
			if request("stepExt") = "04" then
				if chkWorkingUser(request("conno")) then
%>
						<script language="javascript">
							alert("Controller not found")
							history.back();
						</script>
<%
						response.end
					end if 
			end if
			'response.write "extdata = "&request("stepExt")
			if request("stepExt") = "01" then
					repNo = getRep(request("no"), request("orgNo") , request("y"), request("picno"))
					if isNull(repNo) then
							'repNo = rsOra2("sempno")
						repNo = request("picno")
					end if
					sqlOra = "update flow set VAPVNO = '"&request("picno")&"' , VREPNO = '"&repNo&"' where nfrmno = '"&request("no")&"' and " _
							& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '05'"
					rsOra2.open sqlOra, connOra
					'response.write "update PIC = "&sqlOra
					for each asid in split(request("asid"), ", ")
						remark = "remark"&asid
						remark = request(remark)
						reason = "reason"&asid
						reason = request(reason)
						statusapp = "accept"&asid
						statusapp = request(statusapp)
						
						sqlOra = "update fxa_tempfaMst set PICNAME = '"&request("picno")&"' , REMARK = '"&oraFormatStr(remark)&"' , REASON = '"&oraFormatStr(reason)&"' , statusapp = '"&statusapp&"' where " _
								& "id = '"&asid&"'"
						rsOra2.open sqlOra, connOra
					
					next
			else
					if request("stepExt") = "04" then
							repNo = getRep(request("no"), request("orgNo") , request("y"), request("conno"))
							if isNull(repNo) then
									'repNo = rsOra2("sempno")
								repNo = request("conno")
							end if
							sqlOra = "update flow set VAPVNO = '"&request("conno")&"' , VREPNO = '"&repNo&"' where nfrmno = '"&request("no")&"' and " _
									& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '06'"
							rsOra2.open sqlOra, connOra
							' find  FORM SEM
							sqlOra = "select headno from sequenceorg where empno = '"&request("conno")&"'"
							rsCon.open sqlOra, connOra
							if Not rsCon.EOF then
								repNo = getRep(request("no"), request("orgNo") , request("y"), rsCon("headno"))
								if isNull(repNo) then
										'repNo = rsOra2("sempno")
									repNo = rsCon("headno")
								end if
								sqlOra = "update flow set VAPVNO = '"&rsCon("headno")&"' , VREPNO = '"&repNo&"' where nfrmno = '"&request("no")&"' and " _
									& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '07'"
								rsOra2.open sqlOra, connOra
								'find DDEM
								
								sqlOra = "select headno from sequenceorg where empno = '"&rsCon("headno")&"' and sposcode1='21'"
								rsCon2.open sqlOra, connOra
								if Not rsCon2.EOF then
									repNo = getRep(request("no"), request("orgNo") , request("y"), rsCon2("headno"))
									if isNull(repNo) then
											'repNo = rsOra2("sempno")
										repNo = rsCon2("headno")
									end if
									sqlOra = "update flow set VAPVNO = '"&rsCon2("headno")&"' , VREPNO = '"&repNo&"' where nfrmno = '"&request("no")&"' and " _
									& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '08'"
									rsOra2.open sqlOra, connOra
									'find DEM
									sqlOra = "select headno from sequenceorg where empno = '"&rsCon2("headno")&"' and sposcode1='20'"
									rsCon2.close
									rsCon2.open sqlOra, connOra
									if NOT rsCon2.EOF then
										repNo = getRep(request("no"), request("orgNo") , request("y"), rsCon2("headno"))
										if isNull(repNo) then
												'repNo = rsOra2("sempno")
											repNo = rsCon2("headno")
										end if
										sqlOra = "update flow set VAPVNO = '"&rsCon2("headno")&"' , VREPNO = '"&repNo&"' where nfrmno = '"&request("no")&"' and " _
										& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '09'"
										rsOra2.open sqlOra, connOra
									else
										sqlOra = "update flow set cstepst = '"&STEP_NOT_USE&"' where nfrmno = '"&request("no")&"' and " _
												& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '09' "
										rsOra2.open sqlOra, connOra
									
									end if
								
								else
									   ' NOT USE STEP FOR CAN't NOT DDEM
									sqlOra = "update flow set cstepst = '"&STEP_NOT_USE&"' where nfrmno = '"&request("no")&"' and " _
										& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '08' "
									rsOra2.open sqlOra, connOra
									'find DEM
									sqlOra = "select headno from sequenceorg where empno = '"&rsCon("headno")&"' and sposcode1='20'"
									rsCon2.close
									rsCon2.open sqlOra, connOra
									if NOT rsCon2.EOF then
										repNo = getRep(request("no"), request("orgNo") , request("y"), rsCon2("headno"))
										if isNull(repNo) then
												'repNo = rsOra2("sempno")
											repNo = rsCon2("headno")
										end if
										sqlOra = "update flow set VAPVNO = '"&rsCon2("headno")&"' , VREPNO = '"&repNo&"' where nfrmno = '"&request("no")&"' and " _
										& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '09'"
										rsOra2.open sqlOra, connOra
									else
										sqlOra = "update flow set cstepst = '"&STEP_NOT_USE&"' where nfrmno = '"&request("no")&"' and " _
												& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cExtdata = '09' "
										rsOra2.open sqlOra, connOra
									
									end if
								end if
								rsCon2.close
							else
								' NOT USE STEP FOR CAN't NOT HEADNO
								sqlOra = "update flow set cstepst = '"&STEP_NOT_USE&"' where nfrmno = '"&request("no")&"' and " _
									& "vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and (cExtdata = '07' or cExtdata = '08' or cExtdata = '09') "
								rsOra2.open sqlOra, connOra
							end if
							rsCon.close
							'Delete isn't use step
							sqlOra = "select to_char(nFrmNo) as nFrmNo2, to_char(nRunNo) as nRunNo2, flow.* from flow where nFrmNo = '" & request("no") & "' and vOrgNo = '" & request("orgNo") & "' and cYear = '" & request("y") & "' and cYear2 = '" & request("y2") & "' and nRunNo = '" & request("runNo") & "' and cStepSt = '" & STEP_NOT_USE & "'"
							'response.write sqlOra & "<br>"
							rsCon.open sqlOra, connOra
							
							do while not rsCon.EOF
								deleteFlowStep connOra, rsCon("nFrmNo2"), rsCon("vOrgNo"), rsCon("cYear"), rsCon("cYear2"), rsCon("nRunNo2"), "flow", rsCon("cStepNo"), rsCon("cStepNextNo")
								rsCon.moveNext
							loop
							rsCon.close
				
				       end if	
				end if
				Call approve(connOra, request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid") , request("txtRemark"))
				'rsOra.close
				sqlOra = "select cst from form where nfrmno='"&request("no")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"'"
				rsOra.open sqlOra , connOra
				if rsOra("cst") = FLOW_APPROVE then
					sqlOra = "select t.* from fxa_tempfaMst t , fxa_locemp l where t.ID = l.ID and l.nfrmno = '"&request("no")&"' " _
							& "and l.vorgno = '"&request("orgNo") _
							&"' and l.cyear='"&request("y") _
							&"' and l.cyear2='"&request("y2") _
							&"' and nrunno = '"&request("runNo")&"'" _
							&" and t.statusapp = 'T'"
							rsCon.open sqlOra, connOra
							do while not rsCon.EOF
								sqlOra = "select count(assetno)  rnum from fxa_faMst where assetno='"&rsCon("ASSETNO")&"'"
								rsCon2.open sqlOra, connOra
								if rsCon2("rnum") = 0 then
									sqlOra =  "insert into fxa_faMst values('"&rsCon("GRPCODE")&"','"&rsCon("LOCCODE")&"','"&rsCon("ASSETNO")&"','"&oraFormatStr(rsCon("ASSDESC"))&"','"&rsCon("DOCDATE")&"','" _
											&rsCon("INITIALVALUE")&"' , '"&rsCon("STARTDP")&"' ,'"&rsCon("DPRATE")&"' , '"&rsCon("MONTHDP")&"' , '"&rsCon("YTDDP")&"','"&rsCon("ACCUMDP")&"','" _
											&rsCon("BVALUE")&"' , '"&rsCon("INVNO")&"' ,'"&rsCon("MODELNO")&"' , '"&rsCon("SNNO")&"' , '"&rsCon("PONO")&"','"&rsCon("REFASSET")&"','"_
											&rsCon("VOUCHER")&"' , '"&rsCon("QTY")&"' ,'"&rsCon("UNIT")&"' , '"&rsCon("STATUS")&"' , '"&oraFormatStr(rsCon("SUPNAME"))&"','"&rsCon("PRNO")&"','" _
											&rsCon("BUDGETNO")&"' , '"&rsCon("REQBY")&"' , '"&rsCon("ENDDP")&"' , '"&rsCon("USINGLIFT")&"','"&rsCon("PICNAME")&"','" _
											&oraFormatStr(rsCon("REMARK"))&"' , '' , 'N')"
									rsOra2.open sqlOra, connOra
								end if
								rsCon2.close
								rsCon.moveNext
						   loop
						  rsCon.close
				end if
				rsOra.close
		case else
				response.write "Invalid action!"
				response.end
	end select
	set rsOra =nothing
	set rsOra2 =nothing
	set rsCon =nothing
	set rsCon2 =nothing
	set rsCon3 =nothing
connOra.committrans
if (request("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/fixAss/appfaMst.asp"
else
	response.redirect (request("bp"))
end if

%>