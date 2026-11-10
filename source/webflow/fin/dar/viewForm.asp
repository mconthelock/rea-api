<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Disposal Assets Request Form</title>
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
<script language="JavaScript" src="../../js/cntrolFileUpd.js"></script>
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
	Dim totalQty, totalIniValue, cntLine, i
	Dim payto, purpose, paydate, place, description
	Dim ptype, gAMEC, gTotal, sqlSub, totalAmount
	
	set rsOra = createObject("adodb.recordset")
	
	totalQty = 0
	totalIniValue = 0
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

	sqlOra = "select  distinct form.dreqdate, a.sempno inputer, a.sname inpName " _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition, d.writeoffdesc " _
	& " from fxa_darform, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos, fxa_causewriteoff  d  " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and fxa_darform.nfrmno = form.nfrmno  and fxa_darform.vorgno = form.vorgno " _
	& " and fxa_darform.cyear = form.cyear and fxa_darform.cyear2 = form.cyear2 " _
	& " and fxa_darform.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno " _
	& " and fxa_darform.writeoffno = d.writeoffno " _
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
    <div id="templatemo_wrapper">
	<span id="templatemo_wrapper_top"></span>
	        <div id="templatemo_header">    
				<div id="site_title"><h2><b>DISPOSAL ASSETS REQUEST FORM</b></h2></div>
	        </div>       
			<div id="templatemo_main">
			 <div class="col col13 no_margin_right">
					<ul class="tmo_ul_list">
						<li class="first"><strong>Form No : </strong><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),  request.querystring("y2") , request.querystring("runNo"))%></li>
						<li><strong>Input By : </strong><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></li>
						<li><strong>Request By :</strong> <%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></li>
						<li><strong>DIV/DEPT/SEC :</strong> <%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></li>
						<li><strong>Date Request :</strong> <%=rsOra("dreqdate")%></li>
						<li><strong>CAUSE OF WRITE-OFF ASSETS:</strong> <%=left(rsOra("writeoffdesc"),instr(rsOra("writeoffdesc"),"/")-1)%></li>
					</ul>
	        </div>       
	        </div>       			
				<%
				rsOra.close

				sqlOra = "select a.grpcode dgrpcode, a.loccode dloccode, a.assetno dassetno, a.nqty, a.vunit, a.vremark, b.*, c.locname, d.writeoffdesc " _
				& "  from fxa_darform a, fxa_faMst b, fxa_locMst c, fxa_causewriteoff  d " _
				& "  where a.assetno = b.assetno and a.loccode = c.loccode and a.writeoffno = d.writeoffno " _
				& " and a.nfrmno = '" & request.querystring("no") & "'" _
				& " and a.vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and a.cyear = '" & request.querystring("y") & "'" _
				& " and a.cyear2 = '" & request.querystring("y2") & "'" _
				& " and a.nrunno = '" & request.querystring("runNo") & "'"  _
				& " order by a.assetno"
				'response.write sqlOra					
				'response.end
				rsOra.open sqlOra, connOra
				i = 1
				%>
				<table class="table-main" cellspacing="1" cellpadding="1">
					<tr>
						<td style="padding:0px;">
						<table cellspacing="1" cellpadding="1" width="100%">
							<tr>
								<th>ITEM</th><th >GROUP</th><th >STATUS</th><th >ASSET NO.</th><th >L/C CODE</th><th >LOCATION</th><th >DESCRIPTION</th><th >QTY</th>
								<th >PURCHASE DATE</th><th >INITIAL VALUE</th><th >BOOK VALUE</th><!--<th >CAUSE OF WRITE-OFF ASSETS</th>--><th >REMARK</th>
							</tr>
							<% do while not rsOra.EoF	%>
							<tr>
								<td valign=top><%=i%></td>
								<td valign=top><%=rsOra("dgrpcode")%></td>
								<td valign=top><%=rsOra("status")%></td>
								<td valign=top><%=rsOra("dassetno")%></td>
								<td valign=top><%=rsOra("loccode")%></td>
								<td valign=top><%=rsOra("locname")%></td>
								<td valign=top><%=rsOra("assdesc")%></td>
								<td valign=top><%=rsOra("nqty")%></td>
								<td valign=top><%=rsOra("startdp")%></td>
								<td valign=top><%=rsOra("initialvalue")%></td>
								<td valign=top><%=rsOra("bvalue")%></td>
								<!--<td valign=top><%'=left(rsOra("writeoffdesc"),instr(rsOra("writeoffdesc"),"/")-1)%></td>-->
								<td valign=top><%=rsOra("vremark")%>&nbsp;</td>
							</tr>
							
				<%
					'totalQty = totalQty + rsOra("nqty")
					'totalIniValue = totalIniValue + rsOra("initialvalue")					
					i = i + 1 
					rsOra.MoveNext
				loop
				rsOra.close
				%>
							<tr><td colspan=7 align=right><strong>Grand Total</strong></td><td><%=totalQty%></td><td>&nbsp;</td><td><%=totalIniValue%></td><td>&nbsp;</td><td colspan=2>&nbsp;</td></tr>
						</table>
						</td>
					</tr>
				</table>
			
		<div id="templatemo_main">
              <div class="col col13 no_margin_right">
                <ul class="tmo_ul_list">
                    <li class="first"><strong>PHOTOGRAPH : </strong><%
						 call showAttFile(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), "dar","0", mode ,server.urlencode(request.querystring("bp")))
					%></li>
					<li><strong>NOTIFICATION : </strong><%
						 call showAttFile(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), "dar","1", mode ,server.urlencode(request.querystring("bp")))
					%></li>
                    <li><strong>INVESTIGATION :</strong><%
						 call showAttFile(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), "dar","2", mode ,server.urlencode(request.querystring("bp")))
					%></li>
                </ul>
            </div>
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