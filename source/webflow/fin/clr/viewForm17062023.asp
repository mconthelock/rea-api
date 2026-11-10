<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Clear Advance Form</title>
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
		var chk = 0;
		var chk1 = 0;
		if(act == "approve"){
		if (document.getElementById("stepExt").value == "01")
		{
		   var maxid = document.getElementById("maxid").value;
		   
			for(i=1;i<=maxid;i++)
			{
				var acc = "acc"+i;
				var amt = "amt"+i;
				if(document.getElementById(amt).value == "")
				{
					chk1 = 1;
				}
				if(document.getElementById(amt).value != 0)
				{
					if(document.getElementById(acc).value == "")
					{
						chk = 1;
					}
				}
				
			}
			if(chk==1)
			{
				alert("Please input ACCOUNT CODE");
				return false;
			}
			if(chk1==1)
			{
				alert("Please input Amount");
				return false;
			}
		}
		}
		//return false;
		return true;
	}

	function AllowOnlyNumeric()
	{
		var key = window.event.keyCode;
		// Verify if the key entered was a numeric character (0-9)
		if ((key > 47 && key < 58) || (key==46)) return;
		else
			window.event.returnValue = null; 
	}
function currencyFormat (num) {
    return   num.toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, "$1,")
}

function Mycal(){
		
		var maxid = document.getElementById("maxid").value;
		var sump = 0;
			for(i=1;i<=maxid;i++)
			{
			
				var amt = "amt"+i;
				sump = sump + parseFloat(document.getElementById(amt).value);
			
				
			}
		
		document.getElementById("totExp").innerHTML = "<strong>"+currencyFormat (sump)+"</strong>";
		var clfadv = parseFloat(document.getElementById("clfadv").value);
		if (clfadv > sump)
		{
			document.getElementById("txtexp").innerHTML = "<strong>Return cash to company</strong>";
			document.getElementById("exp_adv").innerHTML = "<font color='#FF0000'>("+currencyFormat (clfadv-sump)+")</font>";
			
			
		
		}else{
		
			document.getElementById("txtexp").innerHTML = "<strong>Addition pay to employee</strong>";
			document.getElementById("exp_adv").innerHTML = "<font color='#000000'>"+currencyFormat (sump-clfadv)+"</font>";
			
		}
		
		
	}
</script>	
<%
	Dim rsOra, sqlOra
	Dim mode, requester
	Dim totalAdv, totalExp, cntLine, i
	Dim maxid, htax
	
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
	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') and cstepst='" & STEP_READY & "'"
'	& " and ( vapvno = '" & session("uid") & "' or vrepno = '" & session("uid") & "') "
'	response.write sqlOra
'	response.end
	rsOra.open sqlOra, connOra
	if 	not IsNull(rsOra("cExtdata")) and (not rsOra.EoF) then
		stepExt = rsOra("cExtdata")
	end if
	rsOra.close

	sqlOra = "select  form.dreqdate, a.sempno inputer, a.sname inpName, cyadv, nnoadv, holdtax" _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition , b.sseccode , b.sdepcode , b.sdivcode" _
	& " from clrForm, form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos, cashpayby " _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and clrForm.nfrmno = form.nfrmno  and clrForm.vorgno = form.vorgno " _
	& " and clrForm.cyear = form.cyear and clrForm.cyear2 = form.cyear2 " _
	& " and clrForm.nrunno = form.nrunno and a.sempno = form.vinputer and b.sempno = form.vreqno " _
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
	htax = rsOra("holdtax")
%>	
<div id="templatemo_body_wrapper">
    <div id="templatemo_wrapper"><span id="templatemo_wrapper_top"></span>
        <div id="templatemo_header">    
			<div id="site_title"><h2>Clear Advance Form</h2></div>
        </div>    
	<form name= "frmMain" action="apvForm.asp" method="post" enctype="multipart/form-data" OnSubmit="return CheckData()">
				<% if rsOra("sseccode") <> "00" then 
					if left(rsOra("sseccode"),1) = "0" then
							if left(rsOra("sseccode"),2) = "00" then
				%>
								<input type="hidden" name="costcn" value="<%="10"&right(rsOra("sseccode"),4)%>" />
				<%
							else
							
			%>
								<input type="hidden" name="costcn" value="<%=" "&right(rsOra("sseccode"),5)%>" />
			<%
							end if
					else
			%>
						<input type="hidden" name="costcn" value="<%=rsOra("sseccode")%>" />
			<%
					end if
			%>
				
			<%
				elseif rsOra("sdepcode") <> "00" then
					if left(rsOra("sdepcode"),1) = "0" then
						if left(rsOra("sdepcode"),2) = "00" then
					
			%>
							<input type="hidden" name="costcn" value="<%="10"&right(rsOra("sdepcode"),4)%>" />
			<%
						else
			%>
						<input type="hidden" name="costcn" value="<%=" "&right(rsOra("sdepcode"),5)%>" />
			<%
						end if
				
			
					else
			%>
						<input type="hidden" name="costcn" value="<%=rsOra("sdepcode")%>" />
			<%
					end if
			%>
				
			<%  else 
					if left(rsOra("sdivcode"),1) = "0" then
							if left(rsOra("sdivcode"),2) = "00" then
			%>
								<input type="hidden" name="costcn" value="<%="10"&right(rsOra("sdivcode"),4)%>" />
			<%
							else
			%>
								<input type="hidden" name="costcn" value="<%=" "&right(rsOra("sdivcode"),5)%>" />
			<%
							end if
					else
			%>
							<input type="hidden" name="costcn" value="<%=rsOra("sdivcode")%>" />
		    <%      end if
			%>
					
			<% end if%>
        <div id="templatemo_main">
            <div class="col col13 no_margin_right">
                <ul class="tmo_ul_list">
                    <li class="first"><strong>Form No: </strong><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),  request.querystring("y2") , request.querystring("runNo"))%></li>
					<li><strong>Input By: </strong><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></li>
                    <li><strong>Request By:</strong> <%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></li>
                    <li><strong>DIV/DEPT/SEC:</strong> <%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></li>
                    <li><strong>Date:</strong> <%=rsOra("dreqdate")%></li>
                    <!--<li><strong>Refer PR No:</strong> <%'=rsOra("sprno")%></li>-->
					<% 'if rsOra("payby") <> "" then	%>
                    <!--<li><strong>Pay By:</strong> <%'=rsOra("payby")%></li>
                    <li><strong>Pay Date:</strong> <%'=rsOra("dpaydate")%></li>
                    <li><strong>Pay Time:</strong>&nbsp;<%'=left(rsOra("cpayfrtime"),2)%>:<%'=right(rsOra("cpayfrtime"),2)%>&nbsp;<strong>To</strong>&nbsp;<%'=left(rsOra("cpaytotime"),2)%>:<%'=right(rsOra("cpaytotime"),2)%></li>-->
					<%	'end if	%>
                </ul>
            </div>
			<%
				sqlOra = "select id, description, to_number(namount) namount from cashAdvList where cyear2 = '" & rsOra("cyadv") & "' and nrunno = '" & rsOra("nnoadv") & "'"
				'response.write "sql = "&sqlOra
				rsOra.Close
				rsOra.open sqlOra, connOra				
			%>
			<table width="900" border="1" bgcolor="#fbe9ef" bordercolor="#FFFFFF" cellspacing="2" cellpadding="0">
			<%
				if (stepExt = "01") and (mode=MODE_EDIT) then 
			%>
							
							<input type="hidden" name="clfdes" value="<%=rsOra("description")%>" />
						 <tr bgcolor="#FFFFFF"><td colspan=5><strong>Cash Advance</strong></td></tr>
						  <tr bgcolor="#ec858f" align="center">
							<td><strong>No.</strong></td>
							<td colspan=3><strong>Description</strong></td>
							<td><strong>Amount</strong></td>
						  </tr>
						  <%	do while not rsOra.EoF	%>
						  <tr align="center">
							<td><%=rsOra("id")%></td>
							<td colspan=3><%=rsOra("description")%></td>				
							<td><%=rsOra("namount")%></td>								
						  </tr>
						  <%
										totalAdv = totalAdv + rsOra("namount")
										rsOra.MoveNext	
									loop
									rsOra.close
						%>
						<tr bgcolor="#ec858f">
							<td colspan=4 align=center><strong>Cash Advance Amount</strong></td>
							<td align=center><strong><%=formatNumber(totalAdv, 2, true, true, true)%></strong></td>
						</tr>
						<input type="hidden" name="clfadv" id="clfadv" value="<%=totalAdv%>" />
						<%
							sqlOra = "select id, svendinv, to_number(namount) namount " _
							& " from clrList" _
							& " where clrList.nfrmno = '" & request.querystring("no") & "'" _
							& " and clrList.vorgno = '" & request.querystring("orgNo")  & "'" _
							& " and clrList.cyear = '" & request.querystring("y") & "'" _
							& " and clrList.cyear2 = '" & request.querystring("y2")  & "'" _
							& " and clrList.nrunno = '" & request.querystring("runNo")  & "' "
							
							'response.write sqlOra
							'respone.end
							rsOra.open sqlOra, connOra	
							maxid = 0
						  %>
						  <tr bgcolor="#FFFFFF"><td colspan=5><strong>Clear Advance</strong></td></tr>
						  <tr bgcolor="#ec858f" align="center">
							<td nowrap><strong>No.</strong></td>			  
							<td nowrap><strong>Receipt No</strong></td>
							<td nowrap><strong>ACCOUNT CODE</strong></td>
							<td nowrap><strong>VAT 7%</strong></td>
							<td nowrap><strong>Amount</strong></td>
						  </tr>
							<%	do while not rsOra.EoF 
									maxid = maxid+1
							%>
						  <tr align="center">
							<td nowrap><%=rsOra("id")%></td>
							<td nowrap><%=rsOra("svendinv")%></td>
							<td nowrap><input type="text" name="acc<%=rsOra("id")%>" id="acc<%=rsOra("id")%>" maxlength= "20" class="required input_field" onkeypress = 'return AllowOnlyNumeric()' /></td>
							<td nowrap><input type="checkbox" name="vat<%=rsOra("id")%>" value="1" />
							<td nowrap><input type="text" name="amt<%=rsOra("id")%>" id="amt<%=rsOra("id")%>"  class="required input_field" onkeypress = 'return AllowOnlyNumeric()' value="<%=rsOra("namount")%>" onkeyup = 'Mycal()' /></td>
						  </tr>				
							<%
										'total = total + (rsOra("namount") *rsOra("nvat"))
										totalExp = totalExp + rsOra("namount")
										'cntLine = cntLine + 1
										rsOra.MoveNext
									loop
									rsOra.Close
									cntLine = 1
							
							%>
						<tr bgcolor="#ec858f">
							<td align=center colspan=4><strong>Total Expense Amount</strong></td>
							<td align=center><div id="totExp"><strong><%=formatNumber(totalExp, 2, true, true, true)%></strong></div></td>
						</tr>
						<!--input type="hidden" name="clfamt" value="<%=totalExp%>" /-->
						<input type="hidden" name="maxid" id="maxid" value="<%=maxid%>" />
						<tr>
							<td colspan = 4 align =right><div id="txtexp"><strong>
							<%
								if totalAdv > totalExp then
									color= "#FF0000"
									response.write "Return cash to company"
								else
									color= "#000000"
									response.write "Addition pay to employee"					
								end if
							%>
							</strong></div></td>
							<td align=center><div id="exp_adv">
							<font color=<%=color%>>
							<%=formatNumber(totalExp-totalAdv, 2, true, true, true)%>
							</font></div>
							</td>
						</tr>
						<tr><td colspan = 5><strong>Deduct Withholding tax : </strong><input name="htax" id="htax"  type="text" class="required input_field" onkeypress = 'return AllowOnlyNumeric()' /></td></tr>
			
			<%
				else 
			%>
						 <tr bgcolor="#FFFFFF"><td colspan=3><strong>Cash Advance</strong></td></tr>
						  <tr bgcolor="#ec858f" align="center">
							<td><strong>No.</strong></td>
							<td><strong>Description</strong></td>
							<td><strong>Amount</strong></td>
						  </tr>
						  <%	do while not rsOra.EoF	%>
						  <tr align="center">
							<td><%=rsOra("id")%></td>
							<td><%=rsOra("description")%></td>				
							<td><%=rsOra("namount")%></td>								
						  </tr>
						  <%
										totalAdv = totalAdv + rsOra("namount")
										rsOra.MoveNext	
									loop
									rsOra.close
						%>
						<tr bgcolor="#ec858f">
							<td colspan=2 align=center><strong>Cash Advance Amount</strong></td>
							<td align=center><strong><%=formatNumber(totalAdv, 2, true, true, true)%></strong></td>
						</tr>
						<%
							sqlOra = "select id, svendinv, to_number(namount) namount " _
							& " from clrList" _
							& " where clrList.nfrmno = '" & request.querystring("no") & "'" _
							& " and clrList.vorgno = '" & request.querystring("orgNo")  & "'" _
							& " and clrList.cyear = '" & request.querystring("y") & "'" _
							& " and clrList.cyear2 = '" & request.querystring("y2")  & "'" _
							& " and clrList.nrunno = '" & request.querystring("runNo")  & "' "
							
							'response.write sqlOra
							rsOra.open sqlOra, connOra						
						  %>
						  <tr bgcolor="#FFFFFF"><td colspan=3><strong>Clear Advance</strong></td></tr>
						  <tr bgcolor="#ec858f" align="center">
							<td nowrap><strong>No.</strong></td>			  
							<td nowrap><strong>Receipt No</strong></td>
							<td nowrap><strong>Amount</strong></td>
						  </tr>
							<%	do while not rsOra.EoF %>
						  <tr align="center">
							<td nowrap><%=rsOra("id")%></td>
							<td nowrap><%=rsOra("svendinv")%></td>
							<td nowrap><%=rsOra("namount")%></td>
						  </tr>				
							<%
										'total = total + (rsOra("namount") *rsOra("nvat"))
										totalExp = totalExp + rsOra("namount")
										'cntLine = cntLine + 1
										rsOra.MoveNext
									loop
									rsOra.Close
									cntLine = 1
							%>
						<tr bgcolor="#ec858f">
							<td align=center colspan=2><strong>Total Expense Amount<% %></strong></td>
							<td align=center><strong><%=formatNumber(totalExp, 2, true, true, true)%></strong></td>
						</tr>
						
						<tr>
							<td colspan = 2 align =right><strong>
							<%
								if totalAdv > totalExp then
									color= "#FF0000"
									response.write "Return cash to company"
								else
									color= "#000000"
									response.write "Addition pay to employee"					
								end if
							%>
							</strong></td>
							<td align=center>
							<font color=<%=color%>>
							<%=formatNumber(totalExp-totalAdv, 2, true, true, true)%>
							</font>
							</td>
						</tr>
					   <tr><td colspan = 3><strong>Deduct withholding tax : </strong><%=formatNumber(htax, 2, true, true, true)%></td></tr>
					   
					   
						
						
				<%
					end if
				%>
		</table>
		<div id="contact_form">
		
			<input type="hidden" name="frmNo" value="<%=request.querystring("no")%>">
			<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
			<input type="hidden" name="y" value="<%=request.querystring("y")%>">
			<input type="hidden" name="y2" value="<%=request.querystring("y2")%>">
			<input type="hidden" name="runNo" value="<%=request.querystring("runNo")%>">
			<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
			<input type="hidden" name="requester" value="<%=requester%>">
			<input type="hidden" name="stepExt" id="stepExt" value="<%=stepExt%>">
			<input type="hidden" name="act" value="">
				<table width="900" border="0" cellspacing="2" cellpadding="0">
				<tr>
					<td width=1% nowrap><strong>Attached File:</strong></td>
					<td>
					<%
						 call showAttFile(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), "clr", mode ,server.urlencode(request.querystring("bp")))
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
					<input type="button" name="btnExpExc" value="Export Excel"  class="submit_btn" onclick="javascript:window.location.href='http://<%= Request.ServerVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=FIN_SUB_SYSTEM_FOLDER%>/clr/excelRpt.asp?no=<%=request("no")%>&orgNo=<%=request("orgNo")%>&y=<%=request("y")%>&y2=<%=request("y2")%>&runNo=<%=request("runNo")%>'">																				
					</td>
				</tr>				
				</table>
				
        </div>
    </div>
	</form>		
		
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