<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>FIXED ASSET REGISTER</title>
<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
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

<script language="JavaScript" src="../../js/window.js"></script>
<script language="JavaScript" src="../js/cntrolFileUpd.js"></script>
<script language="JavaScript">
	function CheckData(act)
	{
		if ((act == "approve") && (document.getElementById("stepExt").value=="01")) {
			  var maxs = eval("window.document.frmMain.maxs.value");
			   var i = 0;
			   var chkempty = 0;
			   var chkres = 0 ;
			   for (i = 0; i < maxs; i++){
					var namen = "acceptno"+i;
					var namey = "acceptyes"+i;
					var namerea = "reason"+i;
					
					if (document.getElementById(namen).checked == false && document.getElementById(namey).checked == false)
					{
						chkempty = chkempty + 1; 
					}else{
						
							if (document.getElementById(namen).checked){
								if(document.getElementById(namerea).value == "")
								{
									chkres = chkres + 1;
								}
							}
					}
				}
				
				if (chkempty > 0 ) 
				{
					alert("Please input data");
					return false;
				
				}else{
					if (chkres > 0 ) 
					{
						alert("Plese input reason");
						return false;
					}
				
				}
		}
		return true;
		//return true;
	}
	
function selectrd(v)
{
 var maxs = eval("window.document.frmMain.maxs.value");
	   var i = 0;
	   for (i = 0; i < maxs; i++){
			var namey = "acceptyes"+i;
			var namen = "acceptno"+i;
			  if(v=="T")
			  {
				document.getElementById(namey).checked = true;
				document.getElementById(namen).checked = false;
			  }else{
				document.getElementById(namey).checked = false;
				document.getElementById(namen).checked = true;
			  }
		}
}
</script>	
<%
	Dim rsOra, sqlOra
	Dim mode, requester
	Dim totalAdv, totalExp, cntLine, i
	Dim grp, grpdesc, lc, lcdesc
	Dim rsOra2
	
	set rsOra = createObject("adodb.recordset")
	set rsOra2 = createObject("adodb.recordset")
	
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
'	response.write "test = "&stepExt
	sqlOra = "select  form.dreqdate, a.sempno inputer, a.sname inpName" _
	& " , b.sempno reqNo, b.sname reqName, sec.ssec, dep.sdept, div.sdiv, pos.sposition " _
	& " from form, amec.aemployee a, amec.aemployee b, amec.psection sec, amec.pdepartment dep, amec.pdivision div, amec.pposition pos" _
	& " where b.sseccode = sec.sseccode and b.sdepcode= dep.sdepcode and b.sdivcode = div.sdivcode and b.sposcode = pos.sposcode " _
	& " and a.sempno = form.vinputer and b.sempno = form.vreqno " _
	& " and form.nfrmno = '" & request.querystring("no") & "'" _
	& " and form.vorgno = '" & request.querystring("orgNo")  & "'" _
	& " and form.cyear = '" & request.querystring("y") & "'" _
	& " and form.cyear2 = '" & request.querystring("y2")  & "'" _
	& " and form.nrunno = '" & request.querystring("runNo")  & "'"
'	response.write sqlOra
'	response.end

	rsOra.open sqlOra, connOra
	mode = getMode(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid"))
	'stepExt = getExtdata(request("no"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid")) 
	'response.write "mode = "&mode&"<br/>"
	'response.write "stepExt = "&stepExt&"<br/>"
	
	requester = rsOra("reqNo")
%>	
<div id="templatemo_body_wrapper">
    <div id="templatemo_wrapper"><span id="templatemo_wrapper_top"></span>
        <div id="templatemo_header">    
			<div id="site_title"><h2><b>FIXED ASSET REGISTER FORM</b></h2></div>
        </div>       
        <div id="templatemo_main">
            <div class="col col13 no_margin_right">
                <ul class="tmo_ul_list">
                    <li class="first"><strong>Form No: </strong><%=toFormNumber(request.querystring("no"), request.querystring("orgNo"), request.querystring("y"),  request.querystring("y2") , request.querystring("runNo"))%></li>
					<li><strong>Input By: </strong><%="(" & rsOra("inputer") & ") " & rsOra("inpName") %></li>
                    <li><strong>Request By:</strong> <%="(" & rsOra("reqNo") & ") " & rsOra("reqName") %></li>
                    <li><strong>DIV/DEPT/SEC:</strong> <%=rsOra("sdiv")%>/<%=rsOra("sdept")%>/<%=rsOra("ssec")%></li>
                    <li><strong>Date:</strong> <%=rsOra("dreqdate")%></li>
                </ul>
            </div>
			<%
				sqlOra = "select t.* from fxa_tempfaMst t , fxa_locemp l where t.ID = l.ID and l.nfrmno = '"&request.querystring("no")&"' " _
						& "and l.vorgno = '"&request.querystring("orgNo") _
						&"' and l.cyear='"&request.querystring("y") _
						&"' and l.cyear2='"&request.querystring("y2") _
						&"' and nrunno = '"&request.querystring("runNo")&"'"
				rsOra.Close
				'response.write sqlOra
				rsOra.open sqlOra, connOra	
				if NOT rsOra.EOF then
					grp = rsOra("GRPCODE")
					lc = rsOra("LOCCODE")
					sqlOra = "select ACCASSNAME from fxa_accassMst where ACCASSCODE='"&grp&"'"
					rsOra2.open sqlOra, connOra
					grpdesc = rsOra2("ACCASSNAME")
					rsOra2.close
					sqlOra = "select LOCNAME from fxa_locMst where LOCCODE='"&lc&"'"
					rsOra2.open sqlOra, connOra
					lcdesc = rsOra2("LOCNAME")
					rsOra2.close
				end if
			
				
			%>
			<div id="contact_form">
			<form name="frmMain" method="post" action="approvefa.asp">
				<input type="hidden" name="no" value="<%=request.querystring("no")%>">
				<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
				<input type="hidden" name="y" value="<%=request.querystring("y")%>">
				<input type="hidden" name="y2" value="<%=request.querystring("y2")%>">
				<input type="hidden" name="runNo" value="<%=request.querystring("runNo")%>">
					<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
				<input type="hidden" id="stepExt" name="stepExt" value="<%=stepExt%>">
				
				<input type="hidden" name="act" value=""/>
			<table width="900" border="1" bgcolor="#fbe9ef" bordercolor="#FFFFFF" cellspacing="2" cellpadding="0">
			 <tr bgcolor="#FFFFFF"><td colspan="9"><strong>ASSETS INSPECTION SHEET</strong></td></tr>
			 <tr><td></td></tr>
			 <tr>
				<td colspan="9">
					<table>
					<tr><td colspan="2" width="300"><strong>Group: </strong><%=grp%></td><td colspan="7" width="600"><strong>Group Description: </strong><%=grpdesc%></td></tr>
					<tr><td  colspan="2" width="300"><strong>L/C: </strong><%=lc%></td><td colspan="7" width="600"><strong>Description L/C: </strong><%=lcdesc%></td></tr>
					<% if mode=MODE_EDIT then %>
					<% if stepExt = "01"  then %>
					
							<tr><td  colspan="2" width="400"><strong>PIC(Employee NO.): </strong><input type="text" name="picno" class="input_field" style="width:100px;"/></td><td colspan="7" align="right" valign="bottom" width="500"><a style="font-family: ms sans serif; color: #000000; font-size: 10pt" href="javascript: var winDetail = window.open('detailfa.asp?no=<%=request.querystring("no")%>&orgNo=<%=request.querystring("orgNo")%>&y=<%=request.querystring("y")%>&y2=<%=request.querystring("y2")%>&runNo=<%=request.querystring("runNo")%>', 'winDetail', POPUP_FA_DETAIL); winDetail.focus()">Detail</a></td></tr>
					
					<% 
					else
						if stepExt = "04"  then %>
					
							<tr><td  colspan="2" width="400"><strong>Controller(Employee NO.): </strong><input type="text" name="conno" class="input_field" style="width:100px;"/></td><td colspan="7" align="right" width="500" valign="bottom"><a style="font-family: ms sans serif; color: #000000; font-size: 10pt" href="javascript: var winDetail = window.open('detailfa.asp?no=<%=request.querystring("no")%>&orgNo=<%=request.querystring("orgNo")%>&y=<%=request.querystring("y")%>&y2=<%=request.querystring("y2")%>&runNo=<%=request.querystring("runNo")%>', 'winDetail', POPUP_FA_DETAIL); winDetail.focus()">Detail</a></td></tr>
						<% else %>
							<tr><td  colspan="2" width="400"></td><td colspan="2" width="300"></td><td colspan="7" align="right" width="500" valign="bottom"><a style="font-family: ms sans serif; color: #000000; font-size: 10pt" href="javascript: var winDetail = window.open('detailfa.asp?no=<%=request.querystring("no")%>&orgNo=<%=request.querystring("orgNo")%>&y=<%=request.querystring("y")%>&y2=<%=request.querystring("y2")%>&runNo=<%=request.querystring("runNo")%>', 'winDetail', POPUP_FA_DETAIL); winDetail.focus()">Detail</a></td></tr>
						<% end if %>
					<% end if%>
					
					<% else %>
					<tr><td  colspan="2" width="400"></td><td colspan="7" align="right" width="500" valign="bottom"><a style="font-family: ms sans serif; color: #000000; font-size: 10pt" href="javascript: var winDetail = window.open('detailfa.asp?no=<%=request.querystring("no")%>&orgNo=<%=request.querystring("orgNo")%>&y=<%=request.querystring("y")%>&y2=<%=request.querystring("y2")%>&runNo=<%=request.querystring("runNo")%>', 'winDetail', POPUP_FA_DETAIL); winDetail.focus()">Detail</a></td></tr>
					<% end if %>
					
					</table>
				</td>
			 </tr>
			  <tr bgcolor="#ec858f" align="center">
			  
				<td><strong>ASSET NO.</strong></td>
				<td><strong>Description</strong></td>
				<td><strong>INITIAL VALUE</strong></td>
				<td><strong>QTY</strong></td>
				<td><strong>UNIT</strong></td>
				<td><strong>ACCEPT</strong><% if mode=MODE_EDIT and stepExt = "01" then %><input type="radio" name="accept" value="T" onClick="javascript: selectrd('T')"/><% end if %></td>
				<td><strong>REJECT</strong><% if mode=MODE_EDIT and stepExt = "01" then %><input type="radio" name="accept" value="F" onClick="javascript: selectrd('F')"/><% end if %></td>
				<td><strong>REASON</strong></td>
				<td><strong>REMARK</strong></td>
			  </tr>
			<%
				i=0
				Do While Not rsOra.EOF
						
			%>
			  <tr>
			  <td><%=rsOra("ASSETNO")%><input type="hidden" name="asid" value="<%=rsOra("ID")%>"/><input type="hidden" name="s<%=i%>" value="<%=rsOra("ID")%>"/></td>
			  <td><%=left(rsOra("ASSDESC"),20)%></td>
			  <td align="right"><%=rsOra("INITIALVALUE")%></td>
			  <td align="right"><%=rsOra("QTY")%></td>
			  <td><%=rsOra("UNIT")%></td> 
			  <%
				 if mode=MODE_EDIT and stepExt = "01" then
			  %>
			   <td align="center"><input type="radio" id = "acceptyes<%=i%>" name="accept<%=rsOra("ID")%>" value="T"/></td>
			   <td align="center"><input type="radio" id = "acceptno<%=i%>" name="accept<%=rsOra("ID")%>" value="F"/></td>
			   <td><input type="text" id = "reason<%=i%>" name="reason<%=rsOra("ID")%>" class="input_field" style="width:100px;"/></td>
			  <td><input type="text" id = "remark" name="remark<%=rsOra("ID")%>" class="input_field" style="width:100px;"/></td>
			  <%
				else
			  %>
			    <td align="center"><input type="radio" id = "acceptyes<%=i%>" name="accept<%=rsOra("ID")%>" value="T" disabled="true" <% if rsOra("statusapp") = "T" then Response.write "checked" end if %> /></td>
				<td align="center"><input type="radio" id = "acceptno<%=i%>" name="accept<%=rsOra("ID")%>" value="F" disabled="true" <% if rsOra("statusapp") = "F" then Response.write "checked" end if %> /></td>
				<td><%=rsOra("REASON")%></td>
				<td><%=rsOra("REMARK")%></td>
			  <%
			     end if
			  
			  %>
			  
			  
			  </tr>
			 <% 
				rsOra.MoveNext	
				i = i+1
			  	loop
			 
			 %>
			  <input type="hidden" name="maxs" value="<%=i%>" />
			 </table>
			 <table width="900" border="0" cellspacing="2" cellpadding="0">
				<%	if mode=MODE_EDIT	then	%>
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
					<input type="button" name="btnExpExc" value="Export Excel"  class="submit_btn" onclick="javascript:window.location.href='http://<%= Request.ServerVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=FIN_SUB_SYSTEM_FOLDER%>/fixAss/excelRptfaMst.asp?no=<%=request("no")%>&orgNo=<%=request("orgNo")%>&y=<%=request("y")%>&y2=<%=request("y2")%>&runNo=<%=request("runNo")%>'">
										
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