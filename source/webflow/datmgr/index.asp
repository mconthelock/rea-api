<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<%
	Dim protocol
	If LCase(Request.ServerVariables("HTTPS")) = "on" Then
		protocol = "https://"
	Else
		protocol = "http://"
	End If
%>
<html>
<head> 
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css"></head>
<body> 
<!--#include file="../inc/authen.inc"-->
<table width="100%" border="0" cellpadding="1" cellspacing="0">
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datmgr.gif"></td>
    <td><b><font color="#000000">Data Manager</font></b></td>
  </tr>
<%  if(session("uid") = "13214") or (session("uid") = "13314")  then  %>
    <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/orgLst.asp">Organization Master</a></td>
  </tr>
  
  <% end if %>
  <!--########## SYSTEM MASTER ##########-->
    <%if (getAuthV2(session("sysauth"), AUTH_SYS_ADMIN)) and (session("sysauth") <> "") then%>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">System Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><font color="#000000">Form Master</font></td>
  </tr>	
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0506">IS Forms</a>
	</td>
  </tr>
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=030">GP Forms</a>
	</td>
  </tr>  
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0503">QA Forms</a>
	</td>
  </tr>    
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0510">FE Forms</a>
	</td>
  </tr>    
   <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0504">IE Forms</a>
	</td>
  </tr> 
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0505">PS Forms</a>
	</td>
  </tr>      	
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0801">FIN Forms</a>
	</td>
  </tr>      	
    <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0701">DED Forms</a>
	</td>
  </tr>
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=120">PUR Forms</a>
	</td>
  </tr>  
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=060">MFG Forms</a>
	</td>
  </tr>  
    <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=0702">EPL Forms</a>
	</td>
  </tr>  
  <tr> 
    <td width="1%"></td>
    <td>
	<img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/arrows3.gif" border=0>
		<a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp?ownerNo=110">MAR Forms</a>
	</td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datMgr/rptAuthen.asp">User's 
      authorization</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/stepLst.asp">Step 
      Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/orgLst.asp">Organization Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/orgFrm.asp">
<!--	<a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/orgLst.asp">-->
	Manager Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/leave/holMst.asp">Holiday Master</a></td>
  </tr>  
   <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %>amecweb.mitsubishielevatorasia.co.th/form/isform/IS-RGV/main/setIncharge">Set Incharge Regular Review</a></td>
  </tr>   
  <tr> 
    <td colspan="2"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif" height="4"></td>
  </tr>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Report Manager</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datMgr/rptLst.asp">Report 
      Master</a></td>
  </tr>
  
  <!--########## END SYSTEM MASTER ##########-->
  <%end if%>
  <!--########## OTHER MASTER ##########-->
  
   <%if getAuth("6", "050603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Computer Accessory</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/comAcs/prodLst.asp">Product 
      master</a></td>
  </tr>
  <%end if%>  
  <!--
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/comAcs/pgMstFrm.asp">Product 
      group master</a></td>
  </tr>

  <tr> 
    <td colspan="2"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif" height="4"></td>
  </tr>
  -->  
    <%if getAuth("8", "050603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><font color="#000000"><b>File Server</b> </font></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/fileSvr/svrMstFrm.asp">Server 
      master</a></td>
  </tr>
    <%end if%>  
  <!--
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/fileSvr/fileMstFrm.asp">Directory 
      master</a></td>
  </tr>
	-->    
	<%if getAuth("8", "050603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  	  
  <tr> 
    <td colspan="2"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif" height="4"></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><font color="#000000"><b>E-Mail user requisition</b> </font></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/eMail/respondMstFrm.asp">Set 
      E-Mail to MELINA</a></td>
  </tr>
	<%end if%>
	<%if getAuth("9", "050603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>
  <tr> 
    <td colspan="2"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif" height="4"></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Intranet Content</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/intranet/addMainApv.asp">Main 
      approval master</a></td>
  </tr>
	<%end if%>  
  <!--
	<tr> 
		<td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
		<td><b><font color="#000000">Software Request</font></b></td>
	  </tr>
	  <tr> 
		<td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
		<td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/swReq/grpSW.asp">Software master</a></td>
	  </tr>	-->
  <!--
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><font color="#000000"><b>Budget Requisition</b> </font></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><aa href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/budgetLst.asp">Budget 
      group master</a></td>
  </tr>
  -->
<!--overtime-->
 <%'if getAuth("1", "030201", "03", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
 <%if getAuth("1", "030101", "08", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Bus passenger Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/bus/buslineMst/buslineMst.asp">Bus 
      line master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/bus/busstopMst/busstopMst.asp">bus 
      stop master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/bus/empBusMst/busViewMst.asp">Bus passenger Master</a></td>
  </tr>
    <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/bus/DriverMst/Driver.asp">Driver Master</a></td>
  </tr>
 
    <%end if%>
  <%if getAuth("3", "030101", "16", session("uid"), AUTH_FORM_ADMIN)  or session("uid") = "13314" or session("uid") = "92115" or session("uid") = "14001" or session("uid") = "13202" then%>  
        <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Business Trip master</font></b></td>
  </tr>
	  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/BusTrip/overseaMst.asp">Oversea</a></td>
  </tr>
     <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/BusTrip/currencyMst.asp">Currency</a></td>
  </tr>
    <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/BusTrip/allowancemaster.asp">Allowance Master</a></td>
  </tr>
      <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/TrainTrip/melinaallowancemaster.asp">Melina Allowance Master</a></td>
  </tr>
  
    <%end if%>
 <% 
 if session("uid") = "13314" or session("uid") = "92115" or session("uid") = "16008" or session("uid") = "16007"  then%>  

  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Annual Leave</font></b></td>
  </tr>
     <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/annuallvmst/rightmst.asp">Annual leave right</a></td>
  </tr>
	<tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/annuallvmst/createlv.asp">Create Annual Leave</a></td>
  </tr>
	   <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/annuallvmst/updateemp.asp">Update Employee Card</a></td>
  </tr>
     <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/annuallvmst/setannullv.asp">Upload File Set Annual Leave</a></td>
  </tr>
	 <%end if%>
	

 <%if getAuth("5", "030301", "07", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Document Requisition Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/DocReq/DocReqMst.asp">Requested 
      Document </a></td>
  </tr>
  <%end if%>
   <%if getAuth("1", "030101", "04", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Working performance evaluation</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/gpeva/pgf.asp">Position group master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/gpeva/posl.asp">Position tree master</a></td>
  </tr>
    <%end if%>
 <%if getAuth("10", "030101", "08", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
    <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Stationary</font></b></td>
  </tr> 
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/stnry/prodLst.asp">Stationary master</a></td>
  </tr>  
  <%end if%>
  
 <%if getAuth("2", "030101", "04", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Role structure</font></b></td>
  </tr> 
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/otForm/orgLst.asp">Leader/Foreman Organization master</a></td>
  </tr>  
  <%end if%>
  <%if session("uid")="02035" then%>  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Leave</font></b></td>
  </tr> 
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/Leave/PAupdLv.asp">Update Leave</a></td>
  </tr>    
  <%end if%> 
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Drawing</font></b></td>
  </tr>
   <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %>amecweb.mitsubishielevatorasia.co.th/webflow/datmgr/Masterdwg/index/<%=session("uid")%>">Drawing Request Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="https://amecweb.mitsubishielevatorasia.co.th/webflow/datmgr/dwgnotice/amec/<%=session("uid")%>">AMEC Item Person In-charge</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="https://amecweb.mitsubishielevatorasia.co.th/webflow/datmgr/dwgnotice/subcon/<%=session("uid")%>">Subcon Item Person In-charge</a></td>
  </tr>  
   <!--########## SCRAP MASTER (FOR IBC Dept.[04086: Ms.JARUNEE], PAD Sec., EFC Sec. and P/P Sec.) ##########-->
  <%if session("depcode")="090401" or session("seccode")="090602" or session("seccode")="051002" or session("seccode")="050502" or session("seccode")="050604" then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Scrap</font></b></td>
  </tr>
   <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="https://amecweb.mitsubishielevatorasia.co.th/takeoutsystem/scrap/scrap_master_webflow/<%=session("uid")%>/">Manage scrap product</a></td>
  </tr>
  <%end if%>
  <!--########## END OTHER MASTER ##########-->
   <!--########## AREA MASTER (FOR GA DEPT.) ##########-->
  <%if session("depcode")="040201" or session("depcode")="050601" then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Area to take Photo</font></b></td>
  </tr>
   <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="https://amecweb.mitsubishielevatorasia.co.th/form/gpform/gp-tph/form/area/">Manage Area</a></td>
  </tr>
  <%end if%>
  <!--########## END OTHER MASTER ##########--> 
  <%if getAuth("23", "080101", "26", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Non-PO Master</font></b></td>
  </tr>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="https://amecweb.mitsubishielevatorasia.co.th/form/finform/fin-npo/form/vendor/<%=session("uid")%>/">Vendor for Non-PO Form</a></td>
  </tr>
   <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="https://amecweb.mitsubishielevatorasia.co.th/form/finform/fin-npo/form/ACCCODE/">Account Code</a></td>
  </tr>
  <%end if%>	  
  
  
  <tr> 
    <td width="1%"><img src="<%=protocol %><%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">IS Information</font></b></td>
  </tr>
   <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="<%=protocol %>amecweb.mitsubishielevatorasia.co.th/form/isform/IS-CBO/form/computer/">Check QR Code and Computer Name</a></td>
  </tr>
  
  
    <!--########## END OTHER MASTER ##########-->

</table>
<br>
<!--#include virtual="/sys/inc/copyRight.inc"-->
<!--
<%'if getAuth("5", "030301", "07", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>
  <tr> 
    <td width="1%"><img src="http://<%'=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Budget Requisition Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%'=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%'=request.serverVariables("SERVER_NAME")%>/<%'=SYSTEM_FOLDER%>/<%'=IE_SUB_SYSTEM_FOLDER%>/budget/expMstFrm.asp?bp=<%'=request.serverVariables("SCRIPT_NAME")%>">Expense 
      Type</a></td>
  </tr>
<%'end if%>
	
	
 <%'if getAuth("1", "030101", "04", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>
  <tr> 
    <td width="1%"><img src="http://<%'=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Finger Print Door Access Control Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%'=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%'=request.serverVariables("SERVER_NAME")%>/<%'=SYSTEM_FOLDER%>/<%'=SUB_SYSTEM_FOLDER%>/fingerpIs/fpMstFrm.asp">Room Master</a></td>
  </tr>
    <%'end if%>
-->
</body>
</html>

