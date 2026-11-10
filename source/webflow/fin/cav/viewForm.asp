<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Cash Advance Form</title>
<link href="../css/templatemo_style.css" rel="stylesheet" type="text/css" />

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
<!--
	spe=500;
	na=document.all.tags("blink");
	swi=1;
	bringBackBlinky();
	function bringBackBlinky() {
		if (swi == 1) {
			sho="visible";
			swi=0;
		}else {
			sho="hidden";
			swi=1;
		}
		for(i=0;i<na.length;i++) {
			na[i].style.visibility=sho;
		}
		setTimeout("bringBackBlinky()", spe);
	}
-->	
</script>	
<%
	Dim rsOra, sqlOra
	Dim mode, requester
	Dim total, cntLine, i
	
	set rsOra = createObject("adodb.recordset")
	
	total = 0
	cntLine = 1
	if session("uid") = request.queryString("empNo") then
		apv = request.queryString("empNo")
	else
		apv = session("uid")
	end if		
	
	sqlOra = "select cExtdata from flow where nfrmno = '" & request.querystring("no") & "' and vorgno = '" & request.querystring("orgNo") & "'" _
	& " and cyear = '" & request.querystring("y") & "' and cyear2='" & request.querystring("y2") & "' and nrunno = '" & request.querystring("runNo") & "'" _
	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cExtdata is not null"
'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cstepst='" & STEP_READY & "'"
'	response.write sqlOra
'	response.end
	rsOra.open sqlOra, connOra

	if 	not IsNull(rsOra("cExtdata")) and (not rsOra.EoF) then
		stepExt = rsOra("cExtdata")
	end if
	rsOra.close

	sqlOra = "select  form.dreqdate, cashAdvForm.sprno, a.sempno inputer, a.sname inpName " _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition, cashpayby.payby, cashAdvForm.crecby " _
	& ", to_char(dpaydate,'dd/mm/yyyy') dpaydate, cpayfrtime, cpaytotime, curgent , to_char(effdate,'dd/mm/yyyy') effdate" _
	& " from cashAdvForm, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos, cashpayby " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and cashAdvForm.nfrmno = form.nfrmno  and cashAdvForm.vorgno = form.vorgno " _
	& " and cashAdvForm.cyear = form.cyear and cashAdvForm.cyear2 = form.cyear2 " _
	& " and cashAdvForm.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno and cashpayby.paybyNo(+) = cashAdvForm.cpayby" _
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
			<div id="site_title"><h2>Cash Advance Form</h2></div>
        </div>       
        <div id="templatemo_main">
            <div class="col col13 no_margin_right">
                <ul class="tmo_ul_list">
                    <li class="first"><strong>Form No: </strong><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),  request.querystring("y2") , request.querystring("runNo"))%></li>
                    <li><strong>Input By: </strong><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></li>
                    <li><strong>Request By:</strong> <%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></li>
                    <li><strong>DIV/DEPT/SEC:</strong> <%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></li>
                    <li><strong>Request Date:</strong> <%=rsOra("dreqdate")%></li>
                    <li><strong>Refer PR No:</strong> <%=rsOra("sprno")%></li>
		    <li><blink><font color="red"><strong>Effective Date:</strong>&nbsp;<%=rsOra("effdate")%></font></blink></li>
                    <li><strong>Receive By:</strong>
		    	<%
				if rsOra("crecby") = "0" then
					response.write "Receive cash at couter service FIN sec."
				elseif rsOra("crecby") = "1" then
					response.write "Transfer to Account salary."
				else
					response.write "Transfer to Vendor."
				end if
			%>
		    </li>		    
                    <!--<li><strong>Urgent:</strong><%'if rsOra("curgent") = 0 then response.write "No" else response.write "Yes" end if%></li>					-->
		<% if rsOra("payby") <> "" then	%>
                    <li><strong>Pay By:</strong> <%=rsOra("payby")%></li>
                    <li><strong>Pay Date:</strong> <%=rsOra("dpaydate")%></li>
                    <li><strong>Pay Time:</strong>&nbsp;<%=left(rsOra("cpayfrtime"),2)%>:<%=right(rsOra("cpayfrtime"),2)%>&nbsp;<strong>To</strong>&nbsp;<%=left(rsOra("cpaytotime"),2)%>:<%=right(rsOra("cpaytotime"),2)%></li>
		<%	end if	%>
                </ul>
            </div>
			<%
				rsOra.Close
				sqlOra = "select description, to_number(nquality) nquality, to_number(nprice) nprice, to_number(nvat) as nvat1" _
				& " ,case nvat when 0 then 1 else (nvat/100) end as nvat, (to_number(nquality)* to_number(nprice)) as namount " _
				& " from cashAdvList" _
				& " where cashAdvList.nfrmno = '" & request.querystring("no") & "'" _
				& " and cashAdvList.vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and cashAdvList.cyear = '" & request.querystring("y") & "'" _
				& " and cashAdvList.cyear2 = '" & request.querystring("y2")  & "'" _
				& " and cashAdvList.nrunno = '" & request.querystring("runNo")  & "'"

				sqlOra = "select description, to_number(namount) namount" _
				& " from cashAdvList" _
				& " where cashAdvList.nfrmno = '" & request.querystring("no") & "'" _
				& " and cashAdvList.vorgno = '" & request.querystring("orgNo")  & "'" _
				& " and cashAdvList.cyear = '" & request.querystring("y") & "'" _
				& " and cashAdvList.cyear2 = '" & request.querystring("y2")  & "'" _
				& " and cashAdvList.nrunno = '" & request.querystring("runNo")  & "'"
				
				'response.write sqlOra
				rsOra.open sqlOra, connOra				
			%>
			<p>Purpose for cash advance detail as below: </p>
			<table width="900" border="1" bgcolor="#fbe9ef" bordercolor="#FFFFFF" cellspacing="2" cellpadding="0">
			  <tr bgcolor="#ec858f" align="center">
				<td nowrap><strong>No.</strong></td>			  
				<td nowrap><strong>Description</strong></td>
<!--				<td nowrap><strong>Quantity</strong></td>
				<td nowrap><strong>Price/Unit</strong></td>
				<td nowrap><strong>Vat (%)</strong></td>			
				-->
				<td nowrap><strong>Amount</strong></td>				
			  </tr>
				<%	do while not rsOra.EoF %>
			  <tr align="center">
  				<td nowrap><%=cntLine%></td>
				<td nowrap><%=rsOra("description")%></td>
				<td nowrap><%=formatNumber(rsOra("namount"), 2, true, true, true)%></td>				
<!--				<td nowrap><%'=formatNumber(rsOra("nquality"), 2, true, true, true)%></td>
				<td nowrap><%'=formatNumber(rsOra("nprice"), 2, true, true, true)%></td>
				<td nowrap><%'=formatNumber(rsOra("nvat1"), 2, true, true, true)%></td>			
				<td nowrap><%'=formatNumber((rsOra("namount") *rsOra("nvat")), 2, true, true, true) %></td>							-->
			  </tr>				
				<%
							'total = total + (rsOra("namount") *rsOra("nvat"))
							total = total + rsOra("namount")
							cntLine = cntLine + 1
							rsOra.MoveNext
						loop
						rsOra.Close
						cntLine = 1
				%>
			<tr bgcolor="#ec858f">
				<td colspan=2 align=center><strong>Grand Total</strong></td>
				<td align=center><strong><%=formatNumber(total, 2, true, true, true)%></strong></td>
			</tr>
		</table>
		
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
				<%	if (stepExt = "01") and (mode=MODE_EDIT) then	%>
				<tr>
					<td><strong>Pay By</strong></td>
					<td>
					<%
						sqlOra = "select * from cashpayby order by paybyNo"
						rsOra.open sqlOra, connOra
						do while not rsOra.EoF
						%>
						<input type="radio" name="radcashby" value="<%=rsOra("paybyNo")%>" 
						<%if cntLine = 1 then response.write " checked"%>><%=rsOra("payby")%>
						<%						
								rsOra.MoveNext
								cntLine = cntLine + 1 
							loop
						%>
					</td>
				</tr>
				<tr>
				  <td><strong>Type</strong></td>
				  <td>
						<INPUT TYPE="radio" NAME="radCashTy" VALUE="1" checked>Internal
						<INPUT TYPE="radio" NAME="radCashTy" VALUE="2">External						
				</td>
				</tr>
				<tr>
					<td><strong>Pay Date</strong></td>
					<td>
					<%
						any=TRUE
						call showDateLst("SelDate", day(date+1), "SelMon", month(date+1), "SelYear", year(date+1), year(date)-1, year(date)+1, any, dayExtra, "Onchange='javascript:showDateLst(SelDate, SelMon.value-1,SelYear.value);'", SelYear)
					%>
					</td>
				</tr>
				<tr>
					<td><strong>Pay Time</strong></td>
					<td>
					<select name="selBTimein">
				<% for i=0 to 23 %>
					<option value="<%=i%>" <% if i = 2 then response.write " selected" end if%>><%=chkLenth(i)%></option>										
				<% next	%>
				</select>							
				<select name="selETimein">
					<%	for i=0 to 59%>
							<option value="<%=i%>" <% if i = 0 then response.write " selected" end if%>><%=chkLenth(i)%></option>
					<% next	%>	
				</select> To
				<select name="selBegOut">
					<% for i=0 to 23 %>
						<option value="<%=i%>" <% if i = 17 then response.write " selected" end if%>><%=chkLenth(i)%></option>										
					<% next	%>
					</select>.
					<select name="selEndOut">
						<%	for i=0 to 59%>
								<option value="<%=i%>" <% if i = 30 then response.write " selected" end if%>><%=chkLenth(i)%></option>
						<% next	%>	
				</select>
					</td>
				</tr>				
				<%	end if	%>
				<tr>
					<td width=1% nowrap><strong>Attached File:</strong></td>
					<td>
					<%
						 call showAttFile(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), "cav", mode ,server.urlencode(request.querystring("bp")))
					%>
					</td>
				</tr>
				<%	if (mode=MODE_EDIT)	then	%>
				<tr>
					<td valign=top><input type="button" name="upload" value=" Attach File " onClick="addUpload(0,1500);" class="submit_btn" /></td>
					<td><div id="uploadFile0" name="uploadFile0"> </div></td>
				</tr>				
				<tr>
					<td valign="top"><strong>Remark:</strong></td>
					<td><textarea id="text" name="txtRemark" rows="0" cols="0" class="required"></textarea></td>
				</tr>
				<%	end if	%>				
				<tr>
					<td colspan=2 align=center>
						<%	if (mode=MODE_EDIT)	then	%>
						<input type="button" name="btnApprove" value=" Approve " alt="Approve" title=" Approve " class="submit_btn"  onClick="javascript: if(CheckData('approve')) { document.frmMain.act.value = 'approve'; document.frmMain.submit()}">
						<input type="button" name="btnReject" value=" Reject " alt="Reject" title=" Reject " class="submit_btn"  onClick="javascript: if(CheckData('reject')){document.frmMain.act.value = 'reject'; document.frmMain.submit()}">				
						<%	end if	%>						
						<input type="button" name="btnExpExc" value="Export Excel"  class="submit_btn" onclick="javascript:window.location.href='http://<%= Request.ServerVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=FIN_SUB_SYSTEM_FOLDER%>/cav/excelRpt.asp?no=<%=request("no")%>&orgNo=<%=request("orgNo")%>&y=<%=request("y")%>&y2=<%=request("y2")%>&runNo=<%=request("runNo")%>'">												
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