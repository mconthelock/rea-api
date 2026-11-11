<% Option Explicit %>
<%
           ' Author : Ekkapan Kongchoo
%>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../security/Program.asp" -->

<%
	
    Dim  Rs,sSQL,sSQL2,RS2
	Dim sScriptPath,sexptype
	sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL = "SELECT  * FROM pexpenstype  ORDER BY sexptype "
	sSQL2 = "SELECT  * FROM pVender  ORDER BY sVendcode " 
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 	Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText

%>
<html>
<head>
<title>Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<link rel="stylesheet" href="../styles.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#DB4B37" vlink="#CC3333" alink="#DB4B37" text="#CC3333">
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
    <td width="59%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    <td width="40%" height="2" align="center"><img src="../images/spacer.gif" width="8" height="8"></td>
  </tr>
  <tr>
    <td width="1%">&nbsp;</td>
    <td width="59%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><b><b><b><b><b><b><font color="#BF00BF">Report 
      Manager</font></b></b></b></b></b></b></b></font></td>
    <td width="40%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1%"><img src="../images/spacer.gif" width="1" height="1"></td>
    <td width="59%" align="center" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="98%" height="2" align="left" valign="top" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../images/y01.jpg" width="12" height="12"></b></font></td>
          <td align="right" valign="top" width="2%" height="2">&nbsp;</td>
        </tr>
        <tr> 
          <td width="98%" align="left" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr align="left" valign="top" BGCOLOR="#FFD553"> 
                <td colspan="2" align="center"><img src="../images/okpic006.gif" width="300" height="250"></td>
              </tr>
              <tr align="left" valign="top"> 
                <td width="98%" bgcolor="#FFD553">&nbsp;</td>
                <td width="2%" align="right" bgcolor="#FFD553">&nbsp;</td>
              </tr>
              <tr align="center" valign="top"> 
                <td colspan="2" bgcolor="#FFD553"> 
<%if Mid(Session("sAuthor"),10,3) = "AED" Or Mid(Session("sAuthor"),10,3) = "VIW" then%>
                  <table width="95%" border="0" cellspacing="0" cellpadding="0">
                    <tr bgcolor="#BF00BF"> 
                      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Data 
                        Manager Report</font></b></font></td>
                    </tr>
                    <tr> 
                      <td><img src="../images/plus.gif" width="9" height="9"> 
                        <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pdivisionall.asp">Division 
                        Data</a></font></b> <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pdepartmentall.asp">Department 
                        Data</a></font></b> <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pSectionall.asp">Section 
                        Data</a></font></b> <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pPositionall.asp">Position 
                        Data</a></font></b> <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pexpenstypeall.asp">Expense 
                        Type</a></font></b> <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pCattypeall.asp">Group in Expense 
                        Type</a></font></b> <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pCatalogall.asp">Sub-Group in Expense 
                        Type</a></font></b> <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pproductall.asp">Products</a></font></b> 
                        <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pcurrencyall.asp">Currency</a></font></b> 
                        <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pUnitall.asp">Unit</a></font></b> 
                        <br>
                        <img src="../images/plus.gif" width="9" height="9"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><a href="../report/pvenderall.asp">Vendors</a></font></b> 
                      </td>
                    </tr>
                  </table>
<%End if%>
                </td>
              </tr>
              <tr align="left" valign="top"> 
                <td width="98%" bgcolor="#FFD553">&nbsp;</td>
                <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
              </tr>
            </table>
          </td>
          <td width="2%" align="right"><img src="../images/spacer.gif" width="1" height="1"></td>
        </tr>
      </table>
      <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
    <td width="40%" align="center" valign="top"> 
<%if Mid(Session("sAuthor"),25,3) = "AED" Or Mid(Session("sAuthor"),25,3) = "VIW" then%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr bgcolor="#FFD553"> 
          <td align="left" valign="top" width="2%" height="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFCC66">..</font></b></font></td>
          <td width="96%" height="2" bgcolor="#FFD553">&nbsp;</td>
          <td align="right" valign="top" width="2%" height="2"><img src="../images/y02.jpg" width="12" height="12"></td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%">&nbsp;</td>
          <td width="96%" align="left" valign="top"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#BF00BF"> 
                <td><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><font color="#FFFFFF" size="2">Summary 
                  Report </font></b></font></td>
              </tr>
              <tr valign="top"> 
                <td> 
<!--                  <p><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="SummaryReportOverview.asp"><font size="2" color="#CC00CC">Summary 
                    Report Overview by Expense type</font></a></font></b>
                    </p> 
                    <p><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="SummaryReportByGroup.asp"><font size="2" color="#CC00CC">Summary 
                    Report By Group</font></a></font></b></p>
                    <p><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="SummaryReportByRequesterSection.asp"><font size="2" color="#CC00CC">Summary 
                    Report By Requester Section</font></a></font></b></p> -->
                    <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sGroupReport.asp"><font color="#CC00CC">Report By P/R</font></a></font></b> </p>
                  <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sPoReport.asp"><font color="#CC00CC">Report By P/O</font></a></font></b> </p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sPWhrReport.asp"><font color="#CC00CC">WHR Summary Report</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sPRecReport.asp"><font color="#CC00CC">Report By Receive</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sPInvReport.asp"><font color="#CC00CC">Report By Invoice</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sPTaxReport.asp"><font color="#CC00CC">Report By Tax Invoice</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sOrganizeReport.asp"><font color="#CC00CC">Organization 
                    Report By P/R</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sOrgPoReport.asp"><font color="#CC00CC">Organization 
                    Report By P/O</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sOrgRecReport.asp"><font color="#CC00CC">Organization 
                    Report By Receive</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sOrgInvReport.asp"><font color="#CC00CC">Organization 
                    Report By Invoice</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sOrgTaxReport.asp"><font color="#CC00CC">Organization 
                    Report By Tax Invoice</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sPrcompairReport.asp"><font color="#CC00CC"> 
                    Compair Report By P/R</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/finReport.asp"><font color="#CC00CC"> 
                    Matching Report</font></a></font></b></p>
					<p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/planrecrpt.asp"><font color="#CC00CC">Plan Receive Report</font></a></font></b></p>
					
					
                  <!--########## ADD BY DERMPAN, 18-DEC-2002 ##########-->
                  <hr width="100%" size="1" noshade>
                  <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/poIssueRptQry.asp"><font color="#CC00CC">Purchase 
                    order issue</font></a></font></b>
					                
                  <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/poRcvRptQry.asp"><font color="#CC00CC">Purchase 
                    order received (complete)</font></a></font></b></p> 
					                
                  <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/poRcv2RptQry.asp"><font color="#CC00CC">Purchase 
                    order received (incomplete)</font></a></font></b></p> 
					                
                  <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/outstandingRptQry.asp"><font color="#CC00CC">Purchase 
                    order outstanding</font></a></font></b></p> 
					                
                  <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/poOutSumRptQry.asp"><font color="#CC00CC">Purchase 
                    order outstanding summary</font></a></font></b></p> 
					                
                  <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/poOutDetailRptQry.asp"><font color="#CC00CC">Purchase 
                    order outstanding item list</font></a></font></b> </p>
					
                  <p><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sum_prod.asp"><font color="#CC00CC">Summary item quantity for P/O</font></a></font></b> </p>
                  <hr width="100%" size="1" noshade>
                  <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                  <a href="../report/poItemQryFrm.asp"><font color="#CC00CC">P/O Item</font></a></font></b> 
				  <%if session("sEmpNo") = "system" then%>
                  <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                  <a href="../report/poItemQryFrm.asp"><font color="#CC00CC">Summary P/O Item by P/O</font></a></font></b> 
				  <%end if%>
                  <hr width="100%" size="1" noshade>
                  <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sGroupReportBySubGrp.asp"><font color="#CC00CC">Report by P/R (sub by year)</font></a></font></b><br>
                  <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><img src="../images/plus.gif" width="9" height="9"> 
                    <a href="../report/sPoReportBySubGrp.asp"><font color="#CC00CC">Report by P/O (sub by year)</font></a></font></b>

                  <!--########## END ADD BY DERMPAN, 18-DEC-2002 ##########-->
                </td>
              </tr>
            </table>
            <br>
          </td>
          <td width="2%">&nbsp;</td>
        </tr>
        <tr bgcolor="#FFD553"> 
          <td width="2%" align="left" valign="bottom"><img src="../images/y03.jpg" width="12" height="12"></td>
          <td width="96%" bgcolor="#FFD553" valign="bottom"><img src="../images/spacer.gif" width="1" height="1"></td>
          <td width="2%" align="right" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
        </tr>
      </table>
<%End if%>
    </td>
  </tr>
</table>
<table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr align="center"> 
    <td width="1%"> <img src="../report/images/spacer.gif" width="1" height="1"><br>
    </td>
    <td width="99%"> 
     <!--#include file="../inc/copyRight.inc"--></td>
  </tr>
</table>
</body>
</html>
