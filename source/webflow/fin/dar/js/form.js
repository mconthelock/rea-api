function validate(act)
{
	var i;
	var radioSelected = false;
	//Validate requester
	if(act == "request")
	{
		//check select Evaluation objective
		if ((document.frmMain.radProcAMEC!=undefined) && (document.frmMain.radobj!=undefined)){
		//if (document.frmMain.radProcAMEC[1].checked)  {
			if ((document.frmMain.radobj[0].checked == false) && (document.frmMain.radobj[1].checked == false)) {
				alert("please select Evaluation objective");
				return(false);
			}
		}
		//check input requester		
		if (document.frmMain.txtReqId.value == "" ) {
			alert("please input Requester.");
			document.frmMain.txtReqId.focus();
			return(false);
		}
		//check input Title
		if (document.frmMain.txtTitle.value == "" ) {
			alert("please input Title.");
			document.frmMain.txtTitle.focus();
			return(false);
		}		
		//check input item no		
		if (document.frmMain.txtItemNo.value == "" ) {
			alert("please input Item no.");
			document.frmMain.txtItemNo.focus();
			return(false);
		}
		//check input Drawing No
		if (document.frmMain.txtDwgNo.value == "" ) {
			alert("please input Drawing No.");
			document.frmMain.txtDwgNo.focus();
			return(false);
		}
		//check input Part Name
		if (document.frmMain.txtPrtName.value == "" ) {
			alert("please input Part Name.");
			document.frmMain.txtPrtName.focus();
			return(false);
		}				
		//check input Pur Item No
		if (document.frmMain.txtPurItem.value == "" ) {
			alert("please input Pur Item No.");
			document.frmMain.txtPurItem.focus();
			return(false);
		}				
		//check input PO. or Invoice no.
		if (document.frmMain.txtInvNo.value == "" ) {
			alert("please input PO. or Invoice no");
			document.frmMain.txtInvNo.focus();
			return(false);
		}				
		//check input Order Quantity
		if (document.frmMain.txtOrdQ.value == "" ) {
			alert("please input Order Quantity");
			document.frmMain.txtOrdQ.focus();
			return(false);
		}				
		
		//check input supplier name
		if (document.frmMain.txtSupName.value == "" ) {
			alert("please input Supplier Name.");
			document.frmMain.txtSupName.focus();
			return(false);
		}
		//check input Production Name
		if (document.frmMain.txtPrdName.value == "" ) {
			alert("please input Part Arrive Date / Prod. Month.");
			document.frmMain.txtPrdName.focus();
			return(false);
		}
		//check input Before Changing
		if (document.frmMain.txtBefChg.value == "" ) {
			alert("please input Before Changing.");
			document.frmMain.txtBefChg.focus();
			return(false);
		}			
		//check input After Changing
		if (document.frmMain.txtAftChg.value == "" ) {
			alert("please input After Changing.");
			document.frmMain.txtAftChg.focus();
			return(false);
		}			

		//check input Reason for Others
		if (document.frmMain.radReason[4].checked) {
			if (document.frmMain.txtOther.value=="") {
				alert("please input Reason for others");
				document.frmMain.txtOther.focus();
				return(false);
			}
		}
		
	if (document.frmMain.radLoc[1].checked) {
			if (document.frmMain.txtLoc.value=="") {
				alert("please input Location");
				document.frmMain.txtLoc.focus();
				return(false);
			}
	}
	//check input return
	if (document.frmMain.radSample[1].checked) {
		if (document.frmMain.txtReturn.value=="") {
			alert("please input return to");
			document.frmMain.txtReturn.focus();
			return(false);
		}
	}
	//check input other		
	if (document.frmMain.radSample[2].checked) {
		if (document.frmMain.txtOth.value=="") {
			alert("please input other");
			document.frmMain.txtOth.focus();
			return(false);
		}
	}
//		for (i = 0;  i < frmMain.chkClass.length;  i++)
//		 {
//		    if (frmMain.chkClass[i].checked)
//		        radioSelected = true;
//		 }  
//		 if (!radioSelected)
//		{
//		   alert("Please select classification of changing");
//		   return (false);
//		}	

//		if (frmMain.SelSubmitMon.selectedIndex <= 0)
//		{
//			alert("Please select month.");
//			frmMain.SelSubmitMon.focus();
//			return (false);
//		}
//		if (frmMain.SelSubmitDate.selectedIndex <= 0)
//		{
//			alert("Please select day.");
//			frmMain.SelSubmitDate.focus();
//			return (false);
//		}
//		if (frmMain.SelSubmitYear.selectedIndex <= 0)
//		{
//			alert("Please select year.");
//			frmMain.SelSubmitYear.focus();
//			return (false);
//		}

//		if (frmMain.SelInspMon.selectedIndex <= 0)
//		{
//			alert("Please select month.");
//			frmMain.SelInspMon.focus();
//			return (false);
//		}
//		if (frmMain.SelInspDate.selectedIndex <= 0)
//		{
//			alert("Please select day.");
//			frmMain.SelInspDate.focus();
//			return (false);
//		}
//		if (frmMain.SelInspYear.selectedIndex <= 0)
//		{
//			alert("Please select year.");
//			frmMain.SelInspYear.focus();
//			return (false);
//		}

		if (frmMain.SelExpMon.selectedIndex <= 0)
		{
			alert("Please select month.");
			frmMain.SelExpMon.focus();
			return (false);
		}
		if (frmMain.SelExpDate.selectedIndex <= 0)
		{
			alert("Please select day.");
			frmMain.SelExpDate.focus();
			return (false);
		}
		if (frmMain.SelExpYear.selectedIndex <= 0)
		{
			alert("Please select year.");
			frmMain.SelExpYear.focus();
			return (false);
		}		
	}else if(act == "approve") {
		
	}
	return true;
}

function AllowOnlyNumeric()
{
	var key = window.event.keyCode;
	// Verify if the key entered was a numeric character (0-9)
	if (key > 47 && key < 58)  return;
	else
		window.event.returnValue = null; 
}


//function fixit() {
//  var numberOfElements = document.theForm.elements.length;
//  for (x=0; x<numberOfElements; x++)  {
  // replace all the single, double quotes:
//    var curElement = window.document.theForm.elements[x];
//    curElement.value = curElement.value.replace(/\'/g, "&#39;");
//    curElement.value = curElement.value.replace(/\"/g, "&#34;");
//  }
//  return true;
//}

	  function refreshPanel()
	  {
		setTimeout(refreshPanel,600000);
		my_window  = window.open ("http://webflow/form/cgi-bin/login.asp","test","width=1,height=1,top=1200,left=600,location=no,menuber=no,ienubar=no,toolbar=no,scrollbars=no");
		my_window.close ();
	  }