<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
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

  <!--########## SYSTEM MASTER ##########-->
    <%if getAuthV2(session("sysauth"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">System Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/frmLst.asp">Form 
      Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datMgr/rptAuthen.asp">User's 
      authorization</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/stepLst.asp">Step 
      Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/orgLst.asp">Organization 
      Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datmgr/orgFrm.asp">Manager 
      Master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/leave/holMst.asp">Holiday Master</a></td>
  </tr>  
  <tr> 
    <td colspan="2"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif" height="4"></td>
  </tr>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Report Manager</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/datMgr/rptLst.asp">Report 
      Master</a></td>
  </tr>
  
  <!--########## END SYSTEM MASTER ##########-->
  <%end if%>
  <!--########## OTHER MASTER ##########-->
   <%if getAuth("6", "50603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Computer Accessory</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/comAcs/prodLst.asp">Product 
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
    <%if getAuth("8", "50603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><font color="#000000"><b>File Server</b> </font></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/fileSvr/svrMstFrm.asp">Server master</a></td>
  </tr>
    <%end if%>  
  <!--
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/fileSvr/fileMstFrm.asp">Directory 
      master</a></td>
  </tr>
	-->    
	<%if getAuth("3", "50603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  	  
  <tr> 
    <td colspan="2"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif" height="4"></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><font color="#000000"><b>E-Mail user requisition</b> </font></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=SUB_SYSTEM_FOLDER%>/eMail/respondMstFrm.asp">Set 
      E-Mail to MELINA</a></td>
  </tr>
	<%end if%>
	<%if getAuth("9", "50603", "02", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>
  <tr> 
    <td colspan="2"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif" height="4"></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Intranet Content</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
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

 <%if getAuth("1", "30101", "08", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Bus passenger Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/bus/buslineMst/buslineMst.asp">Bus 
      line master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/bus/busstopMst/busstopMst.asp">bus 
      stop master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/bus/empBusMst/busViewMst.asp">Bus passenger Master</a></td>
  </tr>
    <%end if%>

 <%if getAuth("4", "30101", "08", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Document Requisition Master</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/<%=GP_SUB_SYSTEM_FOLDER%>/DocReq/DocReqMst.asp">Requested 
      Document </a></td>
  </tr>
  <%end if%>
   <%if getAuth("1", "30101", "04", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Working performance evaluation</font></b></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/gpeva/pgf.asp">Position group master</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/gpeva/posl.asp">Position tree master</a></td>
  </tr>
    <%end if%>
 <%if getAuth("10", "30101", "08", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
    <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Stationary</font></b></td>
  </tr> 
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/stnry/prodLst.asp">Stationary master</a></td>
  </tr>  
  <%end if%>
  
 <%'if getAuth("2", "30201", "08", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
 <%if getAuth("2", "30101", "08", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>  
    <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/spacer.gif"></td>
    <td><b><font color="#000000">Role structure</font></b></td>
  </tr> 
  <tr> 
    <td width="1%"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/sys/imgs/datitem.gif"></td>
    <td><a href="/form/gpform/otForm/orgLst.asp">Leader/Foreman Organization master</a></td>
  </tr>  
  <%end if%>

    <!--########## END OTHER MASTER ##########-->
</table>
<br>
<!--#include virtual="/sys/inc/copyRight.inc"-->
<!--
<%'if getAuth("5", "30301", "07", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>
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
	
	
 <%'if getAuth("1", "30101", "04", session("uid"), AUTH_FORM_ADMIN) or getSysAuth(session("uid"), AUTH_SYS_ADMIN) then%>
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

