<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<title>Request For DISPOSAL ASSETS Form</title>
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
			var dataRowHead = "<tr><th>No.</th><th>ASSET NO</th><th>QTY</th><th>UNIT</th><th>Sale Scrap</th></tr>";
			$(".table-persons tbody").empty();
			for(var i = 1; i <= parseInt(valPersons); i++){
					dataRow += "<tr><td align='center'>" + i + ".</td><td><input type='text' name='txtAssDetail" + i + "' id = 'txtAssDetail" + i + "' class='input_field' style='width:150px;'/></td>"; 
					dataRow += "<td><input type='text' name='txtQty" + i + "' id = 'txtQty" + i + "' class='input_field' style='width:50px;'/></td>"; 
					dataRow += "<td><input type='text' name='txtUnit" + i + "' id = 'txtUnit" + i + "' class='input_field' style='width:100px;'/></td>"; 
					//dataRow += "<td><select name='selWriteOff" + i + "' id = 'selWriteOff" + i + "'>";
					//dataRow += "<option value=1>Sale</option>";
					//dataRow += "<option value=2>Donate</option>";
					//dataRow += "<option value=3>Loss</option>";
					//dataRow += "<option value=4>Damage</option>";
					//dataRow += "<option value=5>Other</option>";
					//dataRow += "</select></td>";
					dataRow += "<td><input type='radio' name='radSaleScp" + i + "' id='radSaleScp" + i + "1' value='1' >Yes"; 
					dataRow += "<input type='radio' name='radSaleScp" + i + "' id='radSaleScp" + i + "2' value='2'>No  <input type='text' name='txtRmNo" + i + "' id = 'txtRmNo" + i + "' class='input_field' style='width:200px;'/></td></tr>"; 
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

		var listPart = document.querySelectorAll("[id *= 'txtAssDetail']");
		var listQty = document.querySelectorAll("[id *= 'txtQty']");							
		var listUnit = document.querySelectorAll("[id *= 'txtUnit']");
		var listRmk = document.querySelectorAll("[id *= 'txtRmNo']");			
//		var writOff = document.getElementById("radSaleScp1"+i).value;
//	check input Request by
		if (document.forms["frmMain"]["txtReq"].value == null || document.forms["frmMain"]["txtReq"].value == "") {
			alert("Please input Request by No. ...");
			return false;
		}
//	check input total Asset		
		if (document.forms["frmMain"]["txtTotPersons"].value == null || document.forms["frmMain"]["txtTotPersons"].value == "") {
			alert("Please input total of Asset ...");
			return false;
			
		}else {
				if(listPart.length == 0){
					alert("Please click button OK for add detail asset for disposal");					
					return false;
				}
		}

		if(listPart.length != 0){
			for(var i = 0;i < listPart.length; i++){

				if(listPart[i].value == "" || listPart[i].value == null){
					alert("Please input asset no ...");
					return false;
				}
			}
		}else{
			alert("Please input asset for disposal...");
			return false;
		}
		
		if(listPart.length != 0){
			for(var i = 0;i < listPart.length; i++){
				if(listPart[i].value == "" || listPart[i].value == null){
					alert("Please input asset no ...");
					return false;
				}
			}
		}else{
			alert("Please input asset for disposal...");
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
		
		 if(listUnit.length != 0){
			for(var i = 0;i < listUnit.length; i++){
				if(listUnit[i].value == "" || listUnit[i].value == null){
					alert("Please input unit ...");
					return false;
				}
			}
		}				
				
//		for(var i = 1;i < (listPart.length)+1; i++){
//			var x = document.getElementById("selWriteOff"+i).options.value;					
//			if (x == 4) {	//select case "damage"
//
//				if ((document.getElementById("radSaleScp"+i+"1").checked==false) && (document.getElementById("radSaleScp"+i+"2").checked==false))
//				{
//					alert("Please check move to W/H for item" + i);
//					return false;					
//				}
//
//				if ((document.getElementById("radSaleScp"+[i]+"2").checked==true) && (listRmk[i-1].value==""))
//				{
//					//alert(listRmk[i].value);
//					alert("Please input reason for item " + i);
//					return false;
//				}	
//			} //end x==4
//		} // end for

		return true;					
	}

</script>
</head>
<body>
<!--#include file="../../inc/date.inc"--> 
<!--#include file="../../inc/date2.inc"-->
<div id="templatemo_body_wrapper">
<div id="templatemo_wrapper">
      <div id="templatemo_header"> 
	    <div id="site_title"><h2>DISPOSAL ASSETS REQUEST FORM</h2></div>
	  </div>
	  <div id="templatemo_main">
		 <div id="contact_form">
			<form method="post" name="frmMain" action="apvForm.asp" enctype="multipart/form-data" onsubmit="return validateForm();">
			<input type="hidden" name="frmNo" value="<%=request.querystring("no")%>">
			<input type="hidden" name="orgNo" value="<%=request.querystring("orgNo")%>">
			<input type="hidden" name="y" value="<%=request.querystring("y")%>">
			<input type="hidden" name="bp" value="<%=request.querystring("bp")%>">
			<input type="hidden" name="personsMaxline" value="0"> 
			<input type= "hidden" name = "act" value="request">
			<table border=0 width=100%>
				<tr>
					<td>
						<table>
							<tr>
								<td>
									<b>Request by (Employee No.) : </b><input type="text" name="txtReq" class="input_field" style="width:100px;"/>
								</td>
							</tr>
							<tr>
								<td>
									<b>CAUSE OF WRITE-OFF ASSETS : </b>
									<select name="selWriteOff">
										<option value=1>Sale</option>
										<option value=2>Donate</option>
										<option value=3>Loss</option>
										<option value=4>Damage</option>
										<option value=5>Other</option>
									</select>
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
					</td>
				</tr>
				<tr>
					<td>
						<table class="table-attach">
							<thead>
							<tr>
								<td colspan="3" align="left">
									<b>Attach File (PHOTOGRAPH/√Ÿª∂Ë“¬) (Max 1000 KB) :</b>&nbsp;
									<input type="button" name="upload" value=" Attach File "	onClick="addUpload(0,1000);" class="submit_btn"><div id="uploadFile0" name="uploadFile0"> </div>
								</td>
							</tr>
							</thead>
							<tbody></tbody>
							<thead>
							<tr>
								<td colspan="3" align="left">
									<b>ATTACH (NOTIFICATION/„∫·®Èß§«“¡) (Max 1000 KB)  :</b>&nbsp;
									<input type="button" name="upload" value=" Attach File "	onClick="addUpload(1,1000);" class="submit_btn"><div id="uploadFile1" name="uploadFile1"> </div>
								</td>
							</tr>
							</thead>
							<tbody></tbody>
							<thead>
							<tr>
								<td colspan="3" align="left">
									<b>ATTACH (INVESTIGATION/„∫ Õ∫ «π) (Max 1000 KB) :</b>&nbsp;
									<input type="button" name="upload" value=" Attach File "	onClick="addUpload(2,1000);" class="submit_btn"><div id="uploadFile2" name="uploadFile2"> </div>									
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
									<b>REMARK :</b>&nbsp;
                      <input type="text" name="txtRemark" class="input_field" style="width:200px;" value="">
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