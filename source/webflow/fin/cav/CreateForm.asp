<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Cash Advance Form</title>
<link href="../css/templatemo_style.css" rel="stylesheet" type="text/css" />

<script language="JavaScript" src="../js/cntrolFileUpd.js"></script>
<script language="javascript"  src="../../js/date.js"></script>
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
		newCell.innerHTML = "<center><INPUT TYPE=\"TEXT\" SIZE=\"80\" NAME=\"Column1_"+intLine+"\"  ID=\"Column1_"+intLine+"\" VALUE=\"\"></center>";
 
		//*** Column 2 	Quantity	***//
		//*** Column 2 Amount 	***//
		newCell = newRow.insertCell(1);
		newCell.id = newCell.uniqueID;
		newCell.setAttribute("className", "required input_field");
		newCell.innerHTML = "<center><INPUT TYPE=\"TEXT\"  maxlength=\"10\" SIZE=\"10\" NAME=\"Column2_"+intLine+"\" ID=\"Column2_"+intLine+"\"  VALUE=\"\" onkeypress='AllowOnlyNumeric()' OnChange='chkNum(this)'></center>";

		//*** Column 3 Unit Price 	***//
//		newCell = newRow.insertCell(2);
//		newCell.id = newCell.uniqueID;
//		newCell.setAttribute("className", "required input_field");
//		newCell.innerHTML = "<center><INPUT TYPE=\"TEXT\" maxlength=\"10\" SIZE=\"15\" NAME=\"Column3_"+intLine+"\"  ID=\"Column3_"+intLine+"\" VALUE=\"\" onkeypress='AllowOnlyNumeric()' OnChange='chkNum(this)'></center>";

		//*** Column 4 Vat 	***//
//		newCell = newRow.insertCell(3);
//		newCell.id = newCell.uniqueID;
//		newCell.setAttribute("className", "required input_field");
//		newCell.innerHTML = "<center><INPUT TYPE=\"TEXT\" maxlength=\"3\" SIZE=\"7\" NAME=\"Column4_"+intLine+"\"  ID=\"Column4_"+intLine+"\" VALUE=\"\" onkeypress='AllowOnlyNumeric()' OnChange='chkNum(this)'></center>";
		
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
	if (document.frmMain.radBuyPrd[0].checked) {
		if(eval("document.frmMain.txtprno.value")=="")
		{
			alert('Please input Refer PR No');
			eval("document.frmMain.txtprno.focus();")
			return false;
		}
	}
		if(document.frmMain.hdnMaxLine.value==0) 
		{
			alert('Please input Detail in form');
			return false;
		}
	    var amt = 0;
		for(i=1;i<=document.frmMain.hdnMaxLine.value;i++)
		{
			//*** Description ***/
			if(eval("document.frmMain.Column1_"+i+".value")=="")
			{
				alert('Please input Description line.. ' + i);
				eval("document.frmMain.Column1_"+i+".focus();")
				return false;
			}			
			//***  Quatity ***/				
//			if(eval("document.frmMain.Column2_"+i+".value")=="")
//			{
//				alert('Please input Quatity line.. ' + i);
//				eval("document.frmMain.Column2_"+i+".focus();")
//				return false;
//			}		
			//*** Amount ***/								
			if(eval("document.frmMain.Column2_"+i+".value")=="")
			{
				alert('Please input Amount. line.. ' + i);
				eval("document.frmMain.Column2_"+i+".focus();")
				return false;
			}
			amt = amt + parseFloat(eval("document.frmMain.Column2_"+i+".value"));
			//*** Vat ***/								
//			if(eval("document.frmMain.Column4_"+i+".value")=="")
//			{
//				alert('Please input Vat line.. ' + i);
//				eval("document.frmMain.Column4_"+i+".focus();")
//				return false;
//			}					
		}
	//	if(eval("document.frmMain.txtReq.value") != "12225")
	//	{
			if((document.frmMain.radRecBy[0].checked || document.frmMain.radRecBy[1].checked) &&(amt > 50000))
			{
				alert('Amount over 50,000');
				return false;
			}
	//	}
	
	}		
</script> 
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#0000FF" vlink="#FF0000" alink="#999999">
<!--#include file="../../inc/date.inc"--> 
<!--#include file="../../inc/date2.inc"-->
<div id="templatemo_body_wrapper">
  <div id="templatemo_wrapper"><span id="templatemo_wrapper_top" style="width: 984px; height: 6px"></span> 
    <div id="templatemo_header"> 
	  <div id="site_title"><h2>Cash Advance Form</h2></div>
	</div>
	<div id="templatemo_main">
		<div class="col col23">
		  <div id="contact_form">
		    <form method="post" name="frmMain" action="apvForm.asp" enctype="multipart/form-data" OnSubmit="return CheckValidate();">
			<input type="hidden" name="frmNo" value="<%=request.querystring("no")%>">
			<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
			<input type="hidden" name="y" value="<%=request.querystring("y")%>">
			<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
			<input type="hidden" name="radUrgent" value="0">
			<input type="hidden" name="hdnMaxLine" value="0">
			<input type= "hidden" name = "act" value="request">
			<label for="recvby">Receive By:
				<input type="radio" name="radRecBy" value="0" checked>Receive cash at couter service FIN sec.
				<input type="radio" name="radRecBy" value="1">Transfer to Account salary.
				<input type="radio" name="radRecBy" value="2">Transfer to vendor.				
			</label>
				<label for="recvby">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="radio" name="radBuyPrd" value="0" checked>Purchase goods
				<input type="radio" name="radBuyPrd" value="1">Other
			</label>				
			<!--<label for="urgent">Urgent: 
				<input type="radio" name="radUrgent" value="0" checked>No
				<input type="radio" name="radUrgent" value="1">Yes
			</label>-->
			 <table width="900" border="0">
			 	<tr>
					<td width=1% nowrap><label for="author">Emp Code:</label> </td>
					<td width=1% nowrap><label for="pr">Refer PR No.:</label> 
					<td><label for="effdate">Effective Date:</label> 
				</tr>
				<tr>
					<td width=1% nowrap><input type="text" name="txtReq" maxlength= "12" class="required input_field" /></td>
					<td><input type="text" name="txtprno" maxlength= "12" class="required input_field" /></td>
					<td>
					<%
						any=TRUE
						call showDateLst("SelFLDate", request("SelFLDate"), "SelFLMon", request("SelFLMon"), "SelFLYear", year(date), year(date)-1, year(date)+1, any, dayExtra, "Onchange='javascript:showDateLst(SelFLDate, SelFLMon.value-1,SelFLYear.value);'", SelYear)								
					%></td>					
				</tr>
			 </table>
			<div class="cleaner h10"></div>
			<input type="button" name="upload" value=" Attach File " onClick="addUpload(0,1500);" class="submit_btn" />
			<div id="uploadFile0" name="uploadFile0"> </div>			
			<table width="900" border="0">
				<tr>
					<td width=1% nowrap><label for="text">Purpose for cash advance detail as below: </label></td>
					<td colspan = 3 valign =top align=left><input name="btnAdd" type="button" id="btnAdd" value=" Add " onClick="CreateNewRow();" class="submit_btn"></td>
				</tr>
			</table>
		<table width="900" border="0" id= "tbExp" bgcolor="#fbe9ef">
		  <tr bgcolor="#ec858f" align="center">
			<td nowrap><strong>Description</strong></td>
			<td nowrap><strong>Amount</strong></td>
		  </tr>
		</table>
		<table width="900" border="0" bgcolor="#fbe9ef">
		  <tr>
			<td valign="top"><strong>Remark:</strong></td>
			<td><textarea id="text" name="txtRemark" rows="0" cols="0" class="required"></textarea></td>
		  </tr>		
		  <tr>
		    <td colspan=4 align=center>
				<input type="submit" value="Send" id="submit" name="submit" class="submit_btn" />
                <input type="reset" value="Reset" id="reset" name="reset" class="submit_btn" />			
			</td>
		  </tr>
		</table>
		 </form>
		 </div>
	</div>					
	<div id="templatemo_main">
		<div class="col col23">
			<h4><em>Remark</em></h4>
			<p>1. Please attached supporting document detail as below :-</p>
			<ul class="tmo_ul_list">
				<li class="first">The approved purchase requisition in work flow</li>
				<li>Quotation (If any)</li>
				<li>Other document</li>				
			</ul>				
			<p>2. Cash Advance reimbursement shall be implemented in Thai Baht <em>within 7 days</em> after receiving cash or completing business trip.</p>					
			<p>3. Due date payment By E-Mail</p>								
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