/*
Add by Dermpan on 30-12-03 due to the SCM system preparation for implementation.
*/

function readyForSCM()
{
	return true //if(document.frmMain.Print[2].checked) return false; else return true;
}

function sendToSCM()
{
	if(readyForSCM()) 
	{
		document.frmMain.scm.value = '1'; 
		document.frmMain.Print[1].checked = true;

		if(scmCheck()) document.frmMain.submit()
	}
	else
	{
		alert('Can\'t send canceled P/O to Supply Chain Management system');
	}
}

function removeFromSCM()
{
	if(readyForSCM()) 
	{
		if(confirm('Do you want to remove this P/O from Supply Chain Management System?')) 		
		{
			document.frmMain.scm.value = '2'; document.frmMain.submit()
		}
	}
	else
	{
		alert('Can\'t remove canceled P/O to Supply Chain Management system');
	}
}

function resendToSCM()
{
	if(confirm('Do you want to resend this P/O to Supply Chain Management system?')) sendToSCM();
}

function scmCheck(){
	if(document.frmMain.dDelivery != null && document.frmMain.mDelivery != null && document.frmMain.yDelivery != null){
		if(document.frmMain.dDelivery.options[document.frmMain.dDelivery.selectedIndex].value == "" || document.frmMain.mDelivery.options[document.frmMain.mDelivery.selectedIndex].value == "" || document.frmMain.yDelivery.options[document.frmMain.yDelivery.selectedIndex].value == ""){
			alert("Please specific delivery date!");
			return false;
		}
		else
		{
			return true;
		}
	}

	return false;
}