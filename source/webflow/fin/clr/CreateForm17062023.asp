<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Clear Advance Form</title>
<link href="../css/templatemo_style.css" rel="stylesheet" type="text/css" />

<script language="JavaScript" src="../js/cntrolFileUpd.js"></script>
<script language="JavaScript">
	function CreateNewRow()
	{
		var intLine = parseInt(document.frmMain.hdnMaxLine.value);
		intLine++;
			
		var theTable = document.getElementById("tbExp");
		var newRow = theTable.insertRow(theTable.rows.length)
		newRow.id = newRow.uniqueID
 
		var newCell
		
		//*** Column 1		Description ***//
		newCell = newRow.insertCell(0);
		newCell.id = newCell.uniqueID;
		newCell.setAttribute("className", "required input_field");
		newCell.innerHTML = "<center><INPUT TYPE=\"TEXT\" SIZE=\"20\" NAME=\"Column1_"+intLine+"\"  ID=\"Column1_"+intLine+"\" VALUE=\"\"></center>";
		
		//*** Column 2 Amount 	***//
		newCell = newRow.insertCell(1);
		newCell.id = newCell.uniqueID;
		newCell.setAttribute("className", "required input_field");
		newCell.innerHTML = "<center><INPUT TYPE=\"TEXT\" maxlength=\"10\" SIZE=\"10\" NAME=\"Column2_"+intLine+"\" ID=\"Column2_"+intLine+"\"  VALUE=\"\" onkeypress='AllowOnlyNumeric()' OnChange='chkNum(this)'></center>";
 
		document.frmMain.hdnMaxLine.value = intLine;
	}

	function RemoveRow()
	{
		intLine = parseInt(document.frmMain.hdnMaxLine.value);
		if(parseInt(intLine) > 0)
		{
				theTable = document.getElementById("tbExp");				
				theTableBody = theTable.tBodies[0];
				theTableBody.deleteRow(intLine);
				intLine--;
				document.frmMain.hdnMaxLine.value = intLine;
		}	
	}	

	function AllowOnlyNumeric()
	{
		var key = window.event.keyCode;
		// Verify if the key entered was a numeric character (0-9)
		if ((key > 47 && key < 58) || (key==46)) return;
		else
			window.event.returnValue = null; 
	}
	function chkNum(ele)
	{
		var num = parseFloat(ele.value);
		ele.value = num.toFixed(2);
	}

 function CheckValidate()
	{

		if(eval("document.frmMain.txtReq.value")=="")
		{
			alert('Please input Requester');
			eval("document.frmMain.txtReq.focus();")
			return false;
		}			

		if(eval("document.frmMain.txtyadv.value")=="")
		{
			alert('Please input Refer Advance Year');
			eval("document.frmMain.txtyadv.focus();")
			return false;
		}
///		if (document.frmMain.radType[0].checked) {
			if(eval("document.frmMain.txtadvno.value")=="")
			{
				alert('Please input Refer Advance No');
				eval("document.frmMain.txtadvno.focus();")
				return false;
			}
///		}
				
//		if(document.frmMain.hdnMaxLine.value==0) 
//		{
//			alert('Please input Detail in form');
//			return false;
//		}
		
//		for(i=1;i<=document.frmMain.hdnMaxLine.value;i++)
//		{
			//*** Description ***/
//			if(eval("document.frmMain.Column1_"+i+".value")=="")
//			{
//				alert('Please input Description line.. ' + i);
//				eval("document.frmMain.Column1_"+i+".focus();")
//				return false;
//			}			
			//*** Amount ***/								
//			if(eval("document.frmMain.Column2_"+i+".value")=="")
//			{
//				alert('Please input Amount. line.. ' + i);
//				eval("document.frmMain.Column2_"+i+".focus();")
//				return false;
//			}
//		}				
	}		
</script> 
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#FF0000" alink="#999999">
<%
	Dim rsOra, sqlOra
	set rsOra = createObject("adodb.recordset")
%>	
<div id="templatemo_body_wrapper">
  <div id="templatemo_wrapper">
    <div id="templatemo_header"> 
	  <div id="site_title"><h2>Clear/Cancelled Advance Form</h2></div>
	</div>
	<div id="templatemo_main">
		<div class="col col23">
		  <div id="contact_form">
		    <form method="post" name="frmSearch" action="CreateForm.asp" OnSubmit="return CheckValidate();">
			<input type="hidden" name="no" value="<%=request.querystring("no")%>">
			<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
			<input type="hidden" name="y" value="<%=request.querystring("y")%>">
			<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
			<input type= "hidden" name = "act" value="request">
			<!--<p>
			<label for="author">Request To:
			<input type="radio" name="radType" value="1" checked>Clear
			<input type="radio" name="radType" value="2">Cancelled
			</label>
			</p>
			<br>-->			 
			<p>
				<label for="author">Emp Code:</label> 
				<input type="text" name="txtReq" value="<%=request("txtReq")%>" class="required input_field" />
			</p>
			<p>
				<label for="advyear">Refer Cash Advance Year:</label> 
				<input type="text" name="txtyadv" class="required input_field" value="<%=year(date)%>"/>
			</p>			
			<p>
				<label for="advno">Refer Cash Advance No.:</label> 
				<input type="text" name="txtadvno" value="<%=request("txtadvno")%>" class="required input_field" />
			</p>
			<p><label for="advno">&nbsp;</label> <input type="submit" value="Search" id="submit" name="submit" class="submit_btn" /></p>
			<div class="cleaner h10"></div>	
			</form>
	  </div>
	</div>
	<div id="templatemo_main">
		<div class="col col23">
		  <div id="contact_form">			
		<%
			 If request("submit") <> "" then
				sqlOra = "select clrForm.* from clrForm, form " _
				 & " where cyadv = '" & request("txtyadv") & "' and nnoadv = '" & request("txtadvno") & "'" _
				 & " and clrForm.nfrmno = form.nfrmno and clrForm.vorgno = form.vorgno and clrForm.cyear = form.cyear and clrForm.cyear2 = form.cyear2 and clrForm.nrunno = form.nrunno " _
				 & " and form.cst in ('" & FLOW_RUNNING & "','" & FLOW_APPROVE & "')"
				'response.write sqlOra
				rsOra.open sqlOra, connOra
				if not rsOra.EoF then
					response.write "Request Clear Form is repeat. You can't request form." 
					response.end
				end if			 
				rsOra.Close
				
				sqlOra = "select description, to_number(namount) namount from cashAdvList cav, form where cav.cyear2 = '" & request("txtyadv") & "'" _
				& " and cav.nrunno = '" & request("txtadvno") & "' and form.vreqno = '" & request("txtReq") & "'" _
				& " and cav.nfrmno = form.nfrmno and cav.vorgno = form.vorgno and cav.cyear = form.cyear " _
				& " and cav.cyear2 = form.cyear2 and cav.nrunno = form.nrunno  and form.cst = '" & FLOW_APPROVE & "'" 
				'response.write sqlOra
				rsOra.open sqlOra, connOra
				If rsOra.EoF then
					response.write "<font color ='#FF0000'>Please check input runno for Cash Advance Form." 
					response.write "<br>Data not found or Form is running</font>"
					'response.end
				Else
		%>			 		  
			<form method="post" name="frmMain" action="apvForm.asp" enctype="multipart/form-data">
				<input type="hidden" name="frmNo" value="<%=request("no")%>">
				<input type="hidden" name="orgNo" value="<%=request("orgNo")%>">
				<input type="hidden" name="y" value="<%=request("y")%>">
				<input type="hidden" name="bp" value="<%=request("bp")%>">
				<input type="hidden" name="txtyadv" value="<%=request("txtyadv")%>">
				<input type="hidden" name="txtadvno" value="<%=request("txtadvno")%>">
				<input type="hidden" name="txtReq" value="<%=request("txtReq")%>">
				<input type="hidden" name="hdnMaxLine" value="0">
				<input type= "hidden" name = "act" value="request">
					<table width="900" border="1" bgcolor="#fbe9ef" bordercolor="#FFFFFF" cellspacing="2" cellpadding="0">
					  <tr bgcolor="#ec858f" align="center">
						<td nowrap><strong>Description</strong></td>
						<td nowrap><strong>Amount</strong></td>
					  </tr>
					  <%	Do while not rsOra.EoF %>
					  <tr align="center">
						<td width=50%><%=rsOra("description")%></td>			
						<td nowrap><%=rsOra("namount")%></td>
					  </tr>		  
						<%		
								rsOra.MoveNext
							Loop
						  %>
						  <tr>
						  	<td align=right valign=top><strong>Receipt No</strong></td>
							<td>
								<table width="100%" border="0" id= "tbExp" bgcolor="#fbe9ef" bordercolor="#000000">
									<tr align=center><td><strong>Slip No.</strong></td><td><strong>Amount</strong></td></tr>
								</table>
							</td>
						  </tr>
						  <tr>
						  	<td valign=top>
								<input name="btnAdd" type="button" id="btnAdd" value=" Add Receipt slip No." onClick="CreateNewRow();" class="submit_btn">
								<input type="button" name="upload" value=" Attach File " onClick="addUpload(0,1500);" class="submit_btn" />
							</td>
							<td valign=top><div id="uploadFile0" name="uploadFile0"> </div></td>
						  </tr>
				  		  <tr>
							<td valign="top" width=1%><strong>Remark:</strong></td>
							<td>
								<textarea id="text" name="txtRemark" rows="0" cols="0" class="required"></textarea>								
							</td>
						  </tr>		
						  <tr><td colspan=2 align=center><input type="submit" value="Send" id="submit" name="submit" class="submit_btn" /></td></tr>
						</table>
					 </form>
				<%		End if		
						End If		%>
		</div>
	</div>					
	<div id="templatemo_main">
		<div class="col col23">
			<h4><em>Remark</em></h4>
			<p>1. Please attached original "Receipt slip & Tax invoice"</p>
			<p>2. The approved purchase requisition (P/R) in work flow</p>					
			<p>3. Please Clear Receipt for Refund yourself at FIN Sec.</p>
		</div>
		</div> 
    </div>	
  </div>
</div>
    <div id="templatemo_bottom">
        <div class="cleaner"></div>
    </div>

   <div id="templatemo_footer_wrapper">
        <div id="templatemo_footer">
			<!--#include file="../../inc/copyRight.inc"--> 
        </div>
    </div>
</div>

</body>
</html>