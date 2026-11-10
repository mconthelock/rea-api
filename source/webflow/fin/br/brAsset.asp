<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>Request For ASSETS BORROW Form</title>
<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
<link href="../css/templatemo_style.css" rel="stylesheet" type="text/css" />
<style>
	body{
		font-family:verdana;
	}
	.table-persons tbody th, .table-expense tbody th{
		padding-left:2px;
		padding-right:2px;
		border: 1px solid #FA9BC9;
		background-color:#EED2DB;
	}
	.table-persons tbody td, .table-expense tbody td{
		padding-left:0px;
		padding-right:0px;
		border: 1px solid #FA9BC9;
		background-color:#EED2DB;
	}
</style>
<script src="js/jquery.js"></script>
<script language="JavaScript" src="../../js/cntrolFileUpd.js"></script>
<script language="javascript"  src="../../js/date.js"></script>
<script>
	$(document).ready(function(){
		$("#addPersons").click(function(){
			document.frmMain.personsMaxline.value = document.getElementById("ToltalPerson").value;
			var valPersons = document.getElementById("ToltalPerson").value;
			var dataRow = "";
			//var dataRowHead = "<tr><th>No.</th><th>ASSET NO</th><th>QTY</th><th>UNIT</th><th>CAUSE OF WRITE-OFF</th><th>Sale Scrap</th></tr>";
			var dataRowHead = "<tr><th>No.</th><th>ASSET NO</th><th>QTY</th><th>UNIT</th><th>Description</th><th>Remark</th></tr>";
			$(".table-persons tbody").empty();
			for(var i = 1; i <= parseInt(valPersons); i++){
					dataRow += "<tr><td align='center'>" + i + ".</td><td><input type='text' name='txtAss" + i + "' id = 'txtAss" + i + "'  class='input_field' style='width:150px;' onkeyup='showdetail("+i+")'/></td>"; 
					dataRow += "<td><input type='text' name='txtQty" + i + "' id = 'txtQty" + i + "' class='input_field' style='width:50px;' onkeypress = 'return checkNum(event)'/></td>"; 
					//dataRow += "<td><select name='selWriteOff" + i + "' id = 'selWriteOff" + i + "'>";
					//dataRow += "<option value=1>Sale</option>";
					//dataRow += "<option value=2>Donate</option>";
					//dataRow += "<option value=3>Loss</option>";
					//dataRow += "<option value=4>Damage</option>";
					//dataRow += "<option value=5>Other</option>";
					//dataRow += "</select></td>";
					dataRow += "<td><div id = 'txtUnit" + i + "'>&nbsp;</div></td><td style='width:350px;'><div id='txtDetail"+i+"'>&nbsp;</div></td>"; 
					dataRow += "<td><input type='text' name='txtRmNo" + i + "' id = 'txtRmNo" + i + "' class='input_field' style='width:200px;'/></td></tr>"; 
			}
			if(dataRow != ""){
				dataRowHead += dataRow;
				$(".table-persons tbody").append(dataRowHead);
			}
		});

		$("#btnAttach").click(function(){
			var dataRow = "";
			var list = document.querySelectorAll("[id *= 'filUpload']");
			alert(list[0].value);
			alert(list[1].value);			
			var noNext = list.length;
			dataRow = "<tr id='row" + (noNext+1) + "'><td><input type='file' id='filUpload" + (noNext+1) + "' name='filUpload" + (noNext+1) + "' style='width:400px;'/><input type='button' name='Removefile'  id = '" + (noNext+1) + "' class='RemoveFile' value='Remove' /></td></tr>";
			$(".table-attach tbody1").append(dataRow);
		});
		
		$(".RemoveFile").live("click",function(){
			var id = $(this).attr("id");
			$(".table-attach tbody tr#row"+id).remove();
		});
		
	});
	
	
	function checkNum(e) {
		if (e.keyCode > 31 && (e.keyCode < 48 || e.keyCode > 57)){
			return false;
		}
		else{
			return true;
		}		
	}

	function validateForm() {

		var listAss = document.querySelectorAll("[id *= 'txtAss']");
		var listQty = document.querySelectorAll("[id *= 'txtQty']");	
		if(eval("document.frmMain.empReq.value")=="")
		{
			alert('Please input Requester');
			eval("document.frmMain.empReq.focus();")
			return false;
		}			
		if (document.getElementById("actbr1").checked == false && document.getElementById("actbr2").checked == false)
		{
			alert('Please input CAUSE');
			return false;
		}
	//	if(document.getElementById("actbr1").checked == true){
		//	if(document.getElementById("brloc").value =="")
			//{
				//alert('Please input LC CODE');
				//eval("document.frmMain.brloc.focus();")
				//return false;
			
		//	} 
			//if(eval("document.frmMain.bremp.value")=="")
			//{
				//alert('Please input EMP. CODE');
				//eval("document.frmMain.bremp.focus();")
				//return false;
			
		//	} 
		//}
		if(document.getElementById("actbr2").checked == true){
			if(eval("document.frmMain.brlc.value")=="")
			{
				alert('Please input BORROW BY');
				eval("document.frmMain.brlc.focus();")
				return false;
			
			} 
		}
		if(eval("document.frmMain.dt.value")=="")
		{
			alert('Please input Duration');
			eval("document.frmMain.dt.focus();")
			return false;
		}
		
		if(eval("document.frmMain.txtcause.value")=="")
		{
			alert('Please input Cause of Borrow');
			eval("document.frmMain.txtcause.focus();")
			return false;
		}
		
		if (document.forms["frmMain"]["txtTotPersons"].value == null || document.forms["frmMain"]["txtTotPersons"].value == "") {
			alert("Please input total of Asset ...");
			return false;
			
		}else {
				if(listAss.length == 0){
					alert("Please click button OK for add asset for Borrow");					
					return false;
				}
		}	
		
		if(listAss.length != 0){
			for(var i = 0;i < listAss.length; i++){
				if(listAss[i].value == "" || listAss[i].value == null){
					alert("Please input asset no ...");
					return false;
				}
			}
		}else{
			alert("Please input asset for borrow...");
			return false;
		}
		
		 if(listQty.length != 0){
			for(var i = 0;i < listQty.length; i++){
				if(listQty[i].value == "" || listQty[i].value == null){
					alert("Please input Quantity ...");
					return false;
				}
			}
		}
		

	
		
		//enable('brloc');
		//enable('bremp');
		enable('brlc');
	}
function showRow(rowId) {
    document.getElementById(rowId).style.display = "";
}
function hideRow(rowId) {
    document.getElementById(rowId).style.display = "none";
}

function toggle(srow,hrow){
		showRow(srow);
		hideRow(hrow);
		disableAll();
}

function disableAll()
{
	//disable('brloc');
	//disable('bremp');
	disable('brlc');
}

function selectObtion(op)
{
	disableAll()
	if (op == "BO") 
	{
		//enable('brloc');
		//enable('bremp');
		enable('brlc');
	}

}


function disable(eid)
{
document.getElementById(eid).disabled=true;
}
function enable(eid)
{
document.getElementById(eid).disabled=false;
}
function showdetail(i){
		var asset = document.getElementById("txtAss"+i).value;
		//alert(asset);
	if (asset.length == 0) { 
        return;
    } else {
        var xmlhttp = new XMLHttpRequest();
        xmlhttp.onreadystatechange = function() {
			//alert(xmlhttp.readyState + xmlhttp.status);
            if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
               var x = xmlhttp.responseText;
			   //alert(x);
			   var str = x.split("|");
			   
				// alert("l = "+x.length);
			
			   if (str[0] == null)
			   {
					document.getElementById("txtDetail"+i).innerHTML = "&nbsp;";
					document.getElementById("txtUnit"+i).innerHTML = "&nbsp;";
			   }else{
					document.getElementById("txtDetail"+i).innerHTML = str[0];
					document.getElementById("txtUnit"+i).innerHTML = str[1];
			   }
            }
        }
        xmlhttp.open("GET", "getdetail.asp?a=" + asset, true);
        xmlhttp.send();
    }
	
	}
	function checkNum(e) {
		if (e.keyCode > 31 && (e.keyCode < 48 || e.keyCode > 57)){
			return false;
		}
		else{
			return true;
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
</script>
</head>
<body>
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!--#include file="../../inc/date.inc"--> 
<!--#include file="../../inc/date2.inc"-->
<!--#include file="../inc/showLoc.inc"-->
<div id="templatemo_body_wrapper">
  <div id="templatemo_wrapper"><span id="templatemo_wrapper_top" style="width: 984px; height: 6px"></span> 
    <div id="templatemo_header"> 
	  <div id="site_title"><h2>ASSETS BORROW FORM</h2></div>
	</div>
	<div id="templatemo_main">
		<div class="col col23">
		  <div id="contact_form">
			<%	
				dim rs,sql
				set rs=CreateObject("ADODB.Recordset")
			%>
		    <form method="post" name="frmMain" action="apvBrForm.asp" OnSubmit="return validateForm();">
			<input type="hidden" name="frmNo" value="<%=request.querystring("no")%>">
			<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
			<input type="hidden" name="y" value="<%=request.querystring("y")%>">
			<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
			<input type="hidden" name="radUrgent" value="0">
			<input type="hidden" name="personsMaxline" value="0"> 
			<input type= "hidden" name = "act" value="request">
			 <table width="900" border="0">
				<tr><td width="120"><strong>INPUT BY : </strong></td><td width="780"><%="("&session("uid")&")"&session("uName") %></td></tr>
				<tr><td width="120"><strong>REQUEST BY : </strong></td><td width="780"><input type="text" name="empReq" maxlength= "12" class="required input_field" /></td></tr>
				<!--tr><td width="120"><strong>CAUSE : </strong></td><td></td></tr-->
				<tr>
				 <td colspan="2">
						<table>
							<tr>
								<td width="100"><strong>CAUSE : </strong></td><td width="800"><input type="radio" id = "actbr1" name="actbr" value="BI" onclick="selectObtion('BI')">BORROW INSIDE COMPANY </td>
							</tr>
							<!--tr-->
								<!--td colspan="6"-->
									<!--table>
										<tr><td width="130">&nbsp;</td><td>TO <% 'call showLoc("brloc") %></td><td colspan="4">LC CODE</td></tr -->
										<!--tr><td width="130">&nbsp;</td><td>BORROWER <input type="text" id="bremp" name="bremp" maxlength= "12" class="required input_field" disabled="true" />EMP. CODE</td></tr -->
									<!--/table -->
								<!--/td-->
								<% 'rs.close %>
							<!--/tr-->
							<tr>
								<td width="100">&nbsp;</td><td width="800"><input type="radio" id = "actbr2" name="actbr" value="BO"  onclick="selectObtion('BO')">BORROW OUTSIDE COMPANY &nbsp;&nbsp;BORROW BY&nbsp;<input type="text" id="brlc" name="brlc"  class="required input_field" disabled="true"/></td>
							</tr>
					
						</table>
				 </td>
				
				</tr>
				<tr><td width="100"><strong>Duration : </strong></td><td width="780"><input type="text" id="dt" name="dt"  class="required input_field" /></td></tr>
				<tr>
					<td colspan="2" valign="top">
						<table>
							<tr>
								<td valign="top"><strong>CAUSE OF BORROW : </strong></td><td><textarea id="txtcause" name="txtcause" rows="0" cols="0" class="required"></textarea></td>
							</tr>
						</table>
					
					</td>
				</tr>
			 </table>
			<div class="cleaner h10"></div>
			<table class="table-persons">
							<thead>
							<tr>
								<td colspan="2" align="left">
									<b>Asset :</b> Total
									<input type="text" name="txtTotPersons" value="" id="ToltalPerson" class="input_field" style="width:50px;" onkeypress = 'return checkNum(event)' />&nbsp;Items
								</td>
								<td align="left">
									<input type="button" name="btnPersons" id="addPersons" style="width:50px" class="submit_btn" value="OK"/>
								</td>
							</tr>
							</thead>
							<tbody></tbody>
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