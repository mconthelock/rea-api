<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Request For Entertainment Form</title>
<link href="../css/templatemo_style.css" rel="stylesheet" type="text/css" />
<style>
	body{
		font-family:verdana;
	}
	.table-main {
		margin-left:15px;
	}
	.table-main th{
		border: 1px solid write;
		background-color:#FA9BC9;
	}
	.table-main td{
		padding-left:5px;
		padding-right:5px;
		border: 1px solid #FA9BC9;
		background-color:#EED2DB;
	}
</style>
</head>
<body>
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/flowUtil2.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/date2.inc"--> 
<!--#include file="../../inc/date.inc"-->
<!--#include file="../../inc/misc.inc"-->
<!--#include file="inc/showAttFile.inc"-->
<script language="JavaScript" src="../../js/window.js"></script>
<script language="JavaScript" src="../js/cntrolFileUpd.js"></script>
<script language="JavaScript">
	function CheckData(act)
	{
		return true;
	}
	function createRow(id, value){
		document.getElementById(id).innerHTML = document.getElementById(id).innerHTML + value;
	}
</script>	
<%
	Dim rsOra, sqlOra
	Dim mode, requester
	Dim totalAdv, totalExp, cntLine, i
	Dim payto, purpose, paydate, place, description
	Dim ptype, gAMEC, gTotal, sqlSub, totalAmount
	
	set rsOra = createObject("adodb.recordset")
	
	totalAdv = 0
	totalExp = 0
	cntLine = 1
	if session("uid") = request.queryString("empNo") then
		apv = request.queryString("empNo")
	else
		apv = session("uid")
	end if		
	
	sqlOra = "select cExtdata from flow where nfrmno = '" & request.querystring("no") & "' and vorgno = '" & request.querystring("orgNo") & "'" _
	& " and cyear = '" & request.querystring("y") & "' and cyear2='" & request.querystring("y2") & "' and nrunno = '" & request.querystring("runNo") & "'" _
	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') "
'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cstepst='" & STEP_READY & "'"
'	response.write sqlOra
'	response.end
	rsOra.open sqlOra, connOra

	if 	not IsNull(rsOra("cExtdata")) and (not rsOra.EoF) then
		stepExt = rsOra("cExtdata")
	end if
	rsOra.close

	sqlOra = "select  form.dreqdate, a.sempno inputer, a.sname inpName " _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition " _
	& " from rqtEntForm, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and rqtEntForm.nfrmno = form.nfrmno  and rqtEntForm.vorgno = form.vorgno " _
	& " and rqtEntForm.cyear = form.cyear and rqtEntForm.cyear2 = form.cyear2 " _
	& " and rqtEntForm.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno " _
	& " and form.nfrmno = '" & request.querystring("no") & "'" _
	& " and form.vorgno = '" & request.querystring("orgNo")  & "'" _
	& " and form.cyear = '" & request.querystring("y") & "'" _
	& " and form.cyear2 = '" & request.querystring("y2")  & "'" _
	& " and form.nrunno = '" & request.querystring("runNo")  & "'"
'	response.write sqlOra
'	response.end

	rsOra.open sqlOra, connOra
	mode = getMode(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid"))
'	stepExt = getExtdata(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid")) 	
	requester = rsOra("reqNo")
%>	
<div id="templatemo_body_wrapper">
    <div id="templatemo_wrapper"><span id="templatemo_wrapper_top"></span>
        <div id="templatemo_header">    
			<div id="site_title"><h2><b>Request For Entertainment Form</b></h2></div>
        </div>       
        <div id="templatemo_main">
            <div class="col col13 no_margin_right">
                <ul class="tmo_ul_list">
                    <li class="first"><strong>Form No : </strong><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),  request.querystring("y2") , request.querystring("runNo"))%></li>
					<li><strong>Input By : </strong><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></li>
                    <li><strong>Request By :</strong> <%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></li>
                    <li><strong>DIV/DEPT/SEC :</strong> <%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></li>
                    <li><strong>Date Request :</strong> <%=rsOra("dreqdate")%></li>
                </ul>
            </div>
			<div>
				<%
				rsOra.close
				
				sqlOra = "select a.*, b.sname from rqtEntForm a, amec.aemployee b where " _
				& " a.payto = b.sempno" _
				& " and a.nfrmno = '" & request.querystring("no") & "'" _
				& " and a.vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and a.cyear = '" & request.querystring("y") & "'" _
				& " and a.cyear2 = '" & request.querystring("y2") & "'" _
				& " and a.nrunno = '" & request.querystring("runNo") & "'"
				rsOra.open sqlOra, connOra	
				
				payto = "(" & rsOra("payto") & ") " & rsOra("sname")
				purpose = rsOra("purpose")
				paydate = rsOra("paydate")
				place = rsOra("place")
				description = rsOra("description")
				%>
				<table class="table-main" cellspacing="1" cellpadding="1">
					<tr>
						<th align="center">PAY TO </th><td align="left"><%=payto%></td>
					</tr>
					<tr>
						<th align="center">
							PARTICIPANT(s) 
						</th>
						<td style="padding:0px;">
						<table cellspacing="1" cellpadding="1" width="100%">
							<tr>
								<th width="50%">AMEC</th><th width="50%">GUEST</th>
							</tr>
				<%
				rsOra.close
				
				sqlOra = "select count(*) as A from rqtEntPersons where " _
				& " nfrmno = '" & request.querystring("no") & "'" _
				& " and vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and cyear = '" & request.querystring("y") & "'" _
				& " and cyear2 = '" & request.querystring("y2") & "'" _
				& " and nrunno = '" & request.querystring("runNo") & "' group by ptype "
				rsOra.open sqlOra, connOra
				i = 1
				do while not rsOra.EoF
					if i = 1 then 
						gAMEC = rsOra("A") 
					else 
						if gAMEC >= rsOra("A") then
							ptype = "AMEC"
							sqlSub = "asc"
						else 
							ptype = "GUEST"
							sqlSub = "desc"
						end if 
						gTotal = gAMEC + rsOra("A")
					end if 
					i = i + 1
					rsOra.MoveNext
				loop
				rsOra.close
				i = 1
				j = 1
				
				sqlOra = "select * from rqtEntPersons where " _
				& " nfrmno = '" & request.querystring("no") & "'" _
				& " and vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and cyear = '" & request.querystring("y") & "'" _
				& " and cyear2 = '" & request.querystring("y2") & "'" _
				& " and nrunno = '" & request.querystring("runNo") & "' order by ptype " & sqlSub
				rsOra.open sqlOra, connOra
				do while not rsOra.EoF 
					if rsOra("ptype") = ptype then 
						if rsOra("ptype") = "AMEC" then
							response.write "<tr><td id='pAmec" & i & "' align='left' width='50%'>" & i & ".&nbsp;" & rsOra("pname") & "&nbsp;</td><td id='pGuest" & i & "' align='left' width='50%'>" & i & ".&nbsp;</td></tr>"
						else 
							response.write "<tr><td id='pAmec" & i & "' align='left' width='40%'>" & i & ".&nbsp;</td><td id='pGuest" & i & "' align='left' width='40%'>" & i & ".&nbsp;" & rsOra("pname") & "&nbsp;</td></tr>"
						end if 
					else
						if rsOra("ptype") = "AMEC" then
							response.write "<script>createRow('pAmec" & j & "', '" & rsOra("pname") & "');</script>"
							j = j + 1
						else 
							response.write "<script>createRow('pGuest" & j & "', '" & rsOra("pname") &  "');</script>"
							j = j + 1
						end if
					end if 
					i = i + 1 
					rsOra.MoveNext
				loop
				rsOra.close
				%>
							<tr>
								<td colspan="2" align="center"><b>TOTAL</b>&nbsp;&nbsp;<u><%=gTotal%></u>&nbsp;&nbsp;PERSONS</td>
							</tr>
						</table>
						</td>
					</tr>
					<tr>
						<th align="center">PURPOSE</th><td><%=purpose%>&nbsp;</td>
					</tr>
					<tr>
						<th align="center">DATE</th><td><%=paydate%>&nbsp;</td>
					</tr>
					<tr>
						<th align="center">PLACE</th><td><%=place%>&nbsp;</td>
					</tr>
					<tr>
						<th align="center">DESCRIPTION</th><td><%=description%>&nbsp;</td>
					</tr>
					<tr>
						<th align="center">EXPENSE</th>
						<td style="padding:0px;">
							<table cellspacing="1" cellpadding="1" width="100%" style="border:0px;">
								<tr>
									<th>PARTICULARS</th><th>PRICE</th>
								</tr>
								<%
								totalAmount = 0
								sqlOra = "select * from rqtEntList where " _
								& " nfrmno = '" & request.querystring("no") & "'" _
								& " and vorgno = '" & request.querystring("orgNo")  & "'" _
								& " and cyear = '" & request.querystring("y") & "'" _
								& " and cyear2 = '" & request.querystring("y2") & "'" _
								& " and nrunno = '" & request.querystring("runNo") & "' order by id"
								rsOra.open sqlOra, connOra
								do while not rsOra.EoF
									response.write "<tr><td>" & rsOra("description") & "</td><td align='left'>" & rsOra("namount") & "</td></th>"
									totalAmount = totalAmount + CDbl(rsOra("namount"))
									rsOra.MoveNext
								loop
								%>
								<tr>
									<th>TOTAL(BAHT)</th><th align="center"><%=totalAmount%></th>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			</div>
		<div id="contact_form">
			<form name= "frmMain" action="apvForm.asp" method="post" enctype="multipart/form-data" OnSubmit="return CheckData()">
			<input type="hidden" name="frmNo" value="<%=request.querystring("no")%>">
			<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
			<input type="hidden" name="y" value="<%=request.querystring("y")%>">
			<input type="hidden" name="y2" value="<%=request.querystring("y2")%>">
			<input type="hidden" name="runNo" value="<%=request.querystring("runNo")%>">
			<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
			<input type="hidden" name="requester" value="<%=requester%>">
			<input type="hidden" name="stepExt" value="<%=stepExt%>">
			<input type="hidden" name="act" value="">
				<table width="900" border="0" cellspacing="2" cellpadding="0">
				<tr>
					<td width=1% nowrap><strong>Attached File : </strong></td>
					<td>
					<%
						 call showAttFile(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), "ent", mode ,server.urlencode(request.querystring("bp")))
					%>
					</td>
				</tr>
				<%	if (mode=MODE_EDIT)	then	%>
				<!--<tr>
					<td valign=top><input type="button" name="upload" value=" Attach File " onClick="addUpload(0,1500);" class="submit_btn" /></td>
					<td><div id="uploadFile0" name="uploadFile0"> </div></td>
				</tr>-->
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td valign="top"><strong>Remark : </strong></td>
					<td><textarea id="text" name="txtRemark" rows="0" cols="0" class="required"></textarea></td>
				</tr>
				<%	end if	%>
				<tr>
					<td colspan=2 align=center>
					<%	if (mode=MODE_EDIT)	then	%>
						<input type="button" name="btnApprove" value=" Approve " alt="Approve" title=" Approve " class="submit_btn"  onClick="javascript: if(CheckData('approve')) { document.frmMain.act.value = 'approve'; document.frmMain.submit()}">
						<input type="button" name="btnReject" value=" Reject " alt="Reject" title=" Reject " class="submit_btn"  onClick="javascript: if(CheckData('reject')){document.frmMain.act.value = 'reject'; document.frmMain.submit()}">									
						<%	end if	%>					
					<input type="button" name="btnExpExc" value="Export Excel"  class="submit_btn" onclick="javascript:window.location.href='http://<%= Request.ServerVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=FIN_SUB_SYSTEM_FOLDER%>/ent/excelRpt.asp?no=<%=request("no")%>&orgNo=<%=request("orgNo")%>&y=<%=request("y")%>&y2=<%=request("y2")%>&runNo=<%=request("runNo")%>'">
					<!--<input type="button" name="btnExpExc" value="Export Excel"  class="submit_btn" onclick="javascript:window.location.href='http://<%'= Request.ServerVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=FIN_SUB_SYSTEM_FOLDER%>/ent/excelRpt.asp?no=<%=request("no")%>&orgNo=<%=request("orgNo")%>&y=<%=request("y")%>&y2=<%=request("y2")%>&runNo=<%=request("runNo")%>'">-->																				
					</td>
				</tr>				
				</table>
			</form>			
        </div>
    </div>
		
		
	<br>
<%
		call  showFlow(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),request.querystring("y2"), request.querystring("runNo"), requester, session("uid"),"remark,legend")
%>			
    <div id="templatemo_footer_wrapper">
        <div id="templatemo_footer">
            <!--#include file="../../inc/copyRight.inc"--> 
        </div>
    </div>
</div>
</div>
</body>
</html>