<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>Request For Entertainment Form</title>
<meta http-equiv="X-UA-Compatible" content="IE=Edge" />
<link href="../css/templatemo_style.css" rel="stylesheet" type="text/css" />
<style>
	body{
		font-family:verdana;
	}
	.table-persons tbody th, .table-expense tbody th{
		padding-left:5px;
		padding-right:5px;
		border: 1px solid #FA9BC9;
		background-color:#EED2DB;
	}
	.table-persons tbody td, .table-expense tbody td{
		padding-left:5px;
		padding-right:5px;
		border: 1px solid #FA9BC9;
		background-color:#EED2DB;
	}
</style>
<script src="js/jquery.js"></script>
<script language="javascript"  src="../../js/date.js"></script>
<script>
	$(document).ready(function(){
		$("#addPersons").click(function(){
			document.frmMain.personsMaxline.value = document.getElementById("ToltalPerson").value;
			var valPersons = document.getElementById("ToltalPerson").value;
			var dataRow = "";
			var dataRowHead = "<tr><th>No.</th><th>NAME</th><th>Participant Group</th></tr>";
			$(".table-persons tbody").empty();
			for(var i = 1; i <= parseInt(valPersons); i++){
					dataRow += "<tr><td align='center'>" + i + ".</td><td><input type='text' name='namePersons" + i + "' id = 'namePersons" + i + "' class='input_field' style='width:200px;'/></td>"; 
					dataRow += "<td><input type='radio' name='typeGuest" + i + "' value='AMEC' checked >AMEC&nbsp;"; 
					dataRow += "<input type='radio' name='typeGuest" + i + "' value='GUEST'>GUEST</td></tr>"; 
			}
			if(dataRow != ""){
				dataRowHead += dataRow;
				$(".table-persons tbody").append(dataRowHead);
			}
		});
		
		$("#addExpense").click(function(){
			document.frmMain.expenseMaxline.value = document.getElementById("ToltalExpense").value;
			var dataRow = "";
			var dataRowHead = "<tr><th>No.</th><th>PARTICULARS</th><th>PRICE</th></tr>";
			var dataRowtolPrice ="<tr class='tolPrice'><td>&nbsp;</td><td>TOTAL (BAHT)</td><td><input type='text' name='totalAmount' id='totalAmount' class='input_field' readonly /></td></tr>";
			var Expense = document.getElementById("ToltalExpense").value;
			/*$(".table-expense").empty();
			var list = document.querySelectorAll(".noExpense");
			var noNext = list.length;*/
			$(".table-expense tbody").empty();	
			for(var i = 1; i <= parseInt(Expense); i++){	
				dataRow += "<tr><td class='noExpense' align='center'>" + i + ".</td><td><input type='text' name='partExpense" + i + "' class='input_field' id = 'partExpense" + i + "'/></td>"; 
				dataRow += "<td><input type='text' name='priceExpense" + i + "' id = 'priceExpense" + i + "' class='price input_field' onkeypress = 'return runCal(event," + i + ",0)' ";
				dataRow += "onkeyup = 'return runCal(event," + i + ",1)'/></td></tr>";
			}
			if (dataRow != ""){
				dataRow += dataRowtolPrice;
				dataRowHead += dataRow;
				$(".table-expense tbody").append(dataRowHead);				
			}
		}); 
		
		$("#btnAttach").click(function(){
			var dataRow = "";
			var list = document.querySelectorAll("[id *= 'filUpload']");
			var noNext = list.length;
			dataRow = "<tr id='row" + (noNext+1) + "'><td><input type='file' id='filUpload" + (noNext+1) + "' name='filUpload" + (noNext+1) + "' style='width:400px;'/><input type='button' name='Removefile'  id = '" + (noNext+1) + "' class='RemoveFile' value='Remove' /></td></tr>";
			$(".table-attach tbody").append(dataRow);
		});
		
		$(".RemoveFile").live("click",function(){
			var id = $(this).attr("id");
			$(".table-attach tbody tr#row"+id).remove();
		});
		
		$("#txtOtherPurpose").live("click",function(){
			var id = $(this).attr("id");
			$(".table-attach tbody tr#row"+id).remove();
		});
	});
	
	function checkOthers(e){
		var id = $(e).attr("id");
		if(id == "txtOtherPurpose")
			document.getElementById("otherPurpose").checked = true;
		else if(id == "txtOtherDescript")
			document.getElementById("otherDescript").checked = true;	
	}
	
	function checkNum(e) {
		if (e.keyCode > 31 && (e.keyCode < 48 || e.keyCode > 57)){
			return false;
		}
		else{
			return true;
		}		
	}
	function runCal(e, id, k) {
		if (e.keyCode > 31 && (e.keyCode < 48 || e.keyCode > 57) && e.keyCode != 46){	
			return false;
		}
		else{
			if((k == "1" && e.keyCode == 8) || k == "0"){
				CalAmount(e.keyCode, id);	
			}
			return true;
		}		
	}
	function CalAmount(keyEvent, id){
		var list = document.querySelectorAll(".price");
		var total = 0;
		for(var i = 0;i < list.length; i++){
			//alert(keyEvent);
			if((i+1) == id){
				if(isNaN(total + parseFloat(list[i].value + String.fromCharCode(keyEvent))))
					total += 0;
				else
					total += parseFloat(list[i].value + String.fromCharCode(keyEvent));
			}else{
				if(list[i].value != "")
					total += parseFloat(list[i].value);
				else
					total += 0;
			}		
			document.getElementById("totalAmount").value = total;
		}
	}
	function validateForm() {
		var day, month, year;
		day = document.frmMain.SelFLDate.value;
		month = document.frmMain.SelFLMon.value;
		year = document.frmMain.SelFLYear.value;
		alert(day + " " + month + " " + year);
		if (document.forms["frmMain"]["txtReq"].value == null || document.forms["frmMain"]["txtReq"].value == "") {
			alert("Please input Pay By ...");
			return false;
		}
		var listPart = document.querySelectorAll("[id *= 'namePersons']");
		if(listPart.length != 0){
			for(var i = 0;i < listPart.length; i++){
				if(listPart[i].value == "" || listPart[i].value == null){
					alert("Please input Participant name...");
					return false;
				}
			}
		}else{
			alert("Please input Participant...");
			return false;
		}
		if (document.getElementById("otherPurpose").checked == true) {
			if (document.forms["frmMain"]["txtOtherPurpose"].value == null || document.forms["frmMain"]["txtOtherPurpose"].value == "") {
				alert("Please input Purpose ...");
				return false;
			}
		}
		
		if (document.forms["frmMain"]["SelFLDate"].value == null || document.forms["frmMain"]["SelFLDate"].value == "" || document.forms["frmMain"]["SelFLMon"].value == null || document.forms["frmMain"]["SelFLMon"].value == "" || document.forms["frmMain"]["SelFLYear"].value == null || document.forms["frmMain"]["SelFLYear"].value == ""){
			alert("Please input date ...");
			return false;
		}
		
		if (document.forms["frmMain"]["txtPlace"].value == null || document.forms["frmMain"]["txtPlace"].value == "") {
			alert("Please input Place ...");
			return false;
		}
		if (document.getElementById("otherDescript").checked == true) {
			if (document.forms["frmMain"]["txtOtherDescript"].value == null || document.forms["frmMain"]["txtOtherDescript"].value == "") {
				alert("Please input Description ...");
				return false;
			}
		}
		var listpriceExpense = document.querySelectorAll("[id *= 'priceExpense']");
		var listpartExpense = document.querySelectorAll("[id *= 'partExpense']");
		if(listpriceExpense.length != 0){
			for(var i = 0;i < listpriceExpense.length; i++){
				if(listpriceExpense[i].value == "" || listpriceExpense[i].value == null || listpartExpense[i].value == "" || listpartExpense[i].value == null){
					alert("Please input Expense...");
					return false;
				}
			}
		}else{
			alert("Please input Expense...");
			return false;
		}
		
		var listfilUpload = document.querySelectorAll("[id *= 'filUpload']");
		if(listfilUpload.length != 0){
			for(var i = 0;i < listfilUpload.length; i++){
				if(listfilUpload[i].value == "" || listfilUpload[i].value == null){
					alert("Please attach supporting document.");
					return false;
				}
			}
		}else{
			alert("Please attach supporting document.");
			return false;
		}
	}
</script>
</head>
<body>
<!--#include file="../../inc/date.inc"--> 
<!--#include file="../../inc/date2.inc"-->
<div id="templatemo_body_wrapper">
<div id="templatemo_wrapper"><span id="templatemo_wrapper_top" style="width: 984px; height: 6px"></span> 
      <div id="templatemo_header"> 
	    <div id="site_title"><h2>Request For Entertainment Form</h2></div>
	  </div>
	  <div id="templatemo_main">
		 <div id="contact_form">
			<form method="post" name="frmMain" action="apvForm.asp" enctype="multipart/form-data" onsubmit="return validateForm();">
			<input type="hidden" name="frmNo" value="<%=request.querystring("no")%>">
			<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
			<input type="hidden" name="y" value="<%=request.querystring("y")%>">
			<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
			<input type="hidden" name="expenseMaxline" value="0"> 
			<input type="hidden" name="personsMaxline" value="0"> 
			<input type= "hidden" name = "act" value="request">
			<table>
				<tr>
					<td>
						<table>
							<tr>
								<td>
									<b>PAY BY (Employee No.) : </b><input type="text" name="txtReq" class="input_field" style="width:100px;"/>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table class="table-persons">
							<thead>
							<tr>
								<td colspan="2" align="left">
									<b>PARTICIPANT :</b> Total
									<input type="text" name="txtTotPersons" value="" id="ToltalPerson" class="input_field" style="width:50px;" onkeypress = 'return checkNum(event)' />&nbsp;Persons
								</td>
								<td align="left">
									<input type="button" name="btnPersons" id="addPersons" style="width:50px" class="submit_btn" value="OK"/>
								</td>
							</tr>
							</thead>
							<tbody></tbody>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table class="table-purpose">
							<tr>
								<td align="left">
									<b>PURPOSE : </b><input type="radio" name="typePurpose" value="WELCOME" checked >WELCOME
								</td>
								<td align="left">
									 <input type="radio" name="typePurpose" value="FAREWELL">FAREWELL
								</td>
								<td align="left">
									 <input type="radio" name="typePurpose" value="OTHERS" id="otherPurpose" class="radOthers">OTHERS&nbsp;<input type="text" name="txtOtherPurpose" id="txtOtherPurpose" class="input_field" maxlength="255" onclick="checkOthers(this);"/>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table> 
							<tr>
								<td align="left">
									<b>DATE :</b>&nbsp;
									<%
									any=TRUE
									call showDateLst("SelFLDate", request("SelFLDate"), "SelFLMon", request("SelFLMon"), "SelFLYear", year(date), year(date)-1, year(date)+1, any, dayExtra, "Onchange='javascript:showDateLst(SelFLDate, SelFLMon.value-1,SelFLYear.value);'", SelYear)								
									%>
								</td>
							</tr>
							<tr>
								<td align="left">
									<b>PLACE :</b>&nbsp;<input type="text" name="txtPlace" class="input_field" />
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table class="table-description">
							<tr>
								<td align="left">
									<b>DESCRIPTION : </b><input type="radio" name="typeDescript" value="PARTY" checked >PARTY
								</td>
								<td align="left">
									 <input type="radio" name="typeDescript" value="GIFT">GIFT
								</td>
								<td align="left">
									 <input type="radio" name="typeDescript" value="OTHERS" id="otherDescript" class="radOthers">OTHERS&nbsp;<input type="text" name="txtOtherDescript" id="txtOtherDescript" maxlength="255" class="input_field" onclick="checkOthers(this);" />
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table class="table-expense">
							<thead>
								<tr>
									<td colspan="3" align="left">
										<b>TOTAL PARTICULARS : </b>
										<input type="text" name="txtTotExpense" id="ToltalExpense" class="input_field" onkeypress = 'return checkNum(event)' style="width:50px;" /><b>ITEM</b>
										<input type="button" name="btnExpense" id="addExpense" style="width:90px" class="submit_btn" value="Add"/>
									</td>
								</tr>
							</thead>
							<tbody></tbody>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table class="table-attach">
							<thead>
							<tr>
								<td colspan="3" align="left">
									<b>Attach File :</b>&nbsp;<input type="button" name="btnAttach" id="btnAttach" style="width:90px" class="submit_btn" value="Add"/>
								</td>
							</tr>
							</thead>
							<tbody></tbody>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table>
							<tr>
								<td colspan="3">
									<b>REMARK :</b>&nbsp;<input type="text" name="txtRemark" class="input_field" style="width:200px;" />
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table><br>
			<center><input type="submit" name="btnSend" style="width:100px" value="Send" class="submit_btn" /></center>
			</form>
		</div>
	</div>
	<br>Remark :	(1)  The Receipt can refund under the name and address of  Mitsubishi Elevator Asia Co.,Ltd.
	<br>(2)  Please refer  From No.FIN-ENT000001  in receipt slip.
	<br>(3)  Please send the original receipt slip to FIN SEC. counter service your self.
	<br>(4)  If you have question please contact phone no. 1414 (Anothai)
	<br><br>
	<div id="templatemo_footer_wrapper">
		<div id="templatemo_footer">
			<!--#include file="../../inc/copyRight.inc"--> 
		</div>
	</div>
</div>
	<div id="templatemo_bottom">
        <div class="cleaner"></div>
    </div>
	
</div>
</body>
</html>