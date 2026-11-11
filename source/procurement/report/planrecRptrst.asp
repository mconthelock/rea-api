<% Option Explicit %>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file ="../inc/connUtil.inc" -->
<!-- #include file ="../inc/report.inc" -->
<!-- #include file ="../inc/reportData.inc" -->
<!-- #include file ="../inc/misc.inc" -->

<%
Dim rsOra, sqlOra
Dim color
Dim i, j, k, l, m, poNo
Dim whereClause
Dim row
Dim pageSize, absolutePage, currentRecord, recordCount
Dim totalByPo, grandTotal, total, catType, poRunNo, poDate
Dim cur, curTotal, grp, grpTotal, unit, unitTotal
Dim msg
Dim po, isDate, vendor, exp, expGrp, desc
Dim FromDate,ToDate
Dim div, dept , sec

server.scriptTimeOut = 1800 '30 minute
FromDate = request.queryString("from_day") & "/" & request.queryString("from_month") & "/" & request.queryString("from_year")
ToDate = request.queryString("to_day") & "/" & request.queryString("to_month") & "/" & request.queryString("to_year")
div = request.queryString("ListDivision")
dept = request.queryString("ListDepartment")
sec = request.queryString("ListSection")

Color="#99CCFF"

set rsOra = createObject("adodb.recordset")

'Set page size
if request.queryString("ps") = "" or request.queryString("ps") < 0 then
	pageSize = 50
else
	pageSize = int(request.queryString("ps"))
end if

'Set absolute page
if request.queryString("ab") = "" or request.queryString("ab") < 0 then
	absolutePage = 1
else
	absolutePage = int(request.queryString("ab"))
end if
%>

<HTML>
<HEAD>
<TITLE>P/O ISSUED REPORT</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css" type="text/css">
<style>
font {font-family: ms sans serif; font-size: 9pt}
</style>
<script language="javascript">
	function expExc_onclick() {
		window.location.href="expPlanExcell.asp?fromDate=<%=FromDate%>&toDate=<%=ToDate%>&div=<%=div%>&dept=<%=dept%>&sec=<%=sec%>"
	}
</script>
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		
<Form name="frmMain">
  
<table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
        <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    </tr>
    <tr> 
        <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><font color="#BF00BF">Report 
            Manager</font></b></font></td>
    </tr>
    <tr> 
        <td width="99%" align="center" valign="top" height="0"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                    <td width="98%" height="2" align="left" valign="top" bgcolor="#FFD553"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"><b><img src="../images/y01.jpg" width="12" height="12"></b></font></td>
                </tr>
                <tr> 
                    <td width="98%" align="left" valign="top" height="0"> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr align="left" bgcolor="#FFD553"> 
                                <td colspan="2" align="center" valign="top"> 
                                    <table border=0 width="100%" align="Center" cellspacing="1" cellpadding="0" height="18">
                                        <tr bgcolor="#CC00CC"> 
                                            <td colspan="2" height="20"> 
                                                <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">PLAN RECEIVE REPORT</font></font></b></div>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr> 
                                <td colspan="2" bgcolor="silver"> <font color="blue"><b>PLAN 
                                    DATE: </b><%=request.queryString("from_day") & "/" & request.queryString("from_month") & "/" & request.queryString("from_year") & " - " & request.queryString("to_day") & "/" & request.queryString("to_month") & "/" & request.queryString("to_year")%></font></td>
                            </tr>
                            <tr align="center"> 
                                <td colspan="2" bgcolor="#FFD553" valign="top"> 
                                <table border=0 width="100%" align="Center" cellspacing="1" cellpadding="2" height="18">
                                        <tr bgcolor="#CCCCFF"> 
                                            <td height="20" align="center" width="1%"><b><font color="#000000">NO</font></b></td>
											<td height="20" align="center" width="1%"><b><font color="#000000">PLAN RECEIVE DATE</font></b></td>
											<td height="20" align="center" width="1%"><b><font color="#000000">PLAN RECEIVE TIME</font></b></td>
											<td height="20" align="center" width="1%"><b><font color="#000000">PLAN RECEIVE PLACE</font></b></td>
                                            <td height="20" align="center" width="1%"><b><font color="#000000">P/O NO</font></b></td>
                                            <td height="20" align="center" width="1%"><b><font color="#000000">PRODUCT CODE</font></b></td>
                                            <td height="20" width="1%" align="center"><b><font color="#000000">PRODUCT NAME </font></b></td>
                                            <td height="20" width="1%" align="center"><b><font color="#000000">PLAN QTY</font></b></td>
                                            <td height="20" width="1%" align="center"><b><font color="#000000">UNIT</font></b></td>
											<td height="20" width="1%" align="center"><b><font color="#000000">REQUESTER</font></b></td>
											<td height="20" width="1%" align="center"><b><font color="#000000">SECTION</font></b></td>
											<td height="20" width="1%" align="center"><b><font color="#000000">DEPARTMENT</font></b></td>
											<td height="20" width="1%" align="center"><b><font color="#000000">DIVISION</font></b></td>
                                        </tr>
										 <%
										'########## GENERATE WHERE CLAUSE ##########
										whereClause = "ppoitem.spono = ppo.spono and ppoitem.sprodcode = planreceiveitem.sprodcode " _
										& "and planreceive.spono = ppo.spono and planreceive.ID = planreceiveitem.ID and ppo.srequester = aemployee.sempno " _
										& "and aemployee.sseccode = psection.sseccode and aemployee.sdepcode = pdepartment.sdepcode and aemployee.sdivcode = pdivision.sdivcode " _
										& "and ppo.cstatus <> '9' and (ppo.cstatus = '2' or ppo.cstatus = '3') "
										
										'Po number
										if request.queryString("ListDivision") <> "0" then
											whereClause = whereClause & " and aemployee.sdivcode = '" & request.queryString("ListDivision") & "'"
										end if
										
										'Expense type
										if request.queryString("ListDepartment") <> "0" then
											whereClause = whereClause & " and aemployee.sdepcode  = '" & request.queryString("ListDepartment") & "'"
										end if
										
										'Vendor
										if request.queryString("ListSection") <> "0" then
											whereClause = whereClause & " and aemployee.sseccode = '" & request.queryString("ListSection") & "'"
										end if
										
										'Issue date (from)
										if request.queryString("from_day") <> "" then
											whereClause = whereClause & " and precDate >= to_date('" & request.queryString("from_day") & "/" & request.queryString("from_month") & "/" & request.queryString("from_year") & "', 'DD/MM/YYYY')"
										end if

										'Issue date (to)
										if request.queryString("to_day") <> "" then
											whereClause = whereClause & " and precDate <= to_date('" & request.queryString("to_day")  & "/" & request.queryString("to_month") & "/" & request.queryString("to_year") & "', 'DD/MM/YYYY')"
										end if
								'########## END GENERATE WHERE CLAUSE ##########
										
										'########## GOTO START RECORD ##########
										sqlOra = "select precdate , prectime , precplace , planreceive.spono , planreceiveitem.sprodcode , planreceiveitem.sprodname , planreceiveitem.nquality , planreceiveitem.sunitcode" _
										& ", aemployee.sname , psection.ssec , pdepartment.sdept , pdivision.sdiv" _
										& " from ppo, ppoItem, aemployee , planreceive , planreceiveitem , psection , pdepartment , pdivision " _
										& "where " & whereClause _
										& " order by precdate, prectime , planreceive.sPoNo, planreceiveitem.sProdName "
									    'response.write sqlOra
										rsOra.open sqlOra, conn
									
										recordCount = 1
										
										do while not rsOra.EOF and recordCount  < absolutePage * pageSize - pageSize + 1
											recordCount = recordCount + 1
											rsOra.moveNext
										loop
									
										currentRecord = absolutePage * pageSize - pageSize + 1
										row = currentRecord
										
										'########## END GOTO START RECORD ##########
									
										If Not rsOra.EOF then
											i = 0
											poRunNo = 1
											poNo = rsOra("sPoNo")
											j = 0
											l = 0
											m = 0
											
											do while not rsOra.EOF and row < absolutePage * pageSize + 1
												
										%>
                                        <tr bgcolor="#FFCCCC" valign="top"> 
                                            <td width="1%" bgcolor="<%=Color%>" height="2" align="center"><font color="#000000"><%=row%></font></td>
											<td width="1%" bgcolor="<%=Color%>" height="2"> 
                                                <font color="#000000"> 
                                                <%=rsOra("precdate")%>
                                                </font></td>
											<td width="1%" bgcolor="<%=Color%>" height="2"> 
                                                <font color="#000000"> 
                                                <%=rsOra("prectime")%>
                                                </font></td>
											<td width="1%" bgcolor="<%=Color%>" height="2"> 
                                                <font color="#000000"> 
                                                <%=rsOra("precplace")%>
                                                </font></td>
                                            <td width="1%" bgcolor="<%=Color%>" height="2"> 
                                                <font color="#000000"> 
                                                <%=rsOra("spono")%>
                                                </font></td>
											 <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=rsOra("sprodcode")%>
                                                </font></td>
												 <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=rsOra("sProdName")%>
                                                </font></td>
												 <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=FormatNumber(rsOra("nquality"))%>
                                                </font></td>
												 <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=rsOra("sunitcode")%>
                                                </font></td>
											  
												  <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=rsOra("sname")%>
                                                </font></td>
												  <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=rsOra("ssec")%>
                                                </font></td>
													  <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=rsOra("sdept")%>
                                                </font></td>
													  <td height="20" bgcolor="<%=Color%>"> 
                                                <font color="#000000"> 
                                                <%=rsOra("sdiv")%>
                                                </font></td>
												
                               </tr>
                                        <%
											if Color="#FFCCCC" then
												Color="#99CCFF"
											else
												Color="#FFCCCC"
											end if
											
											color = "#99CCFF"
											
											row = row + 1
											rsOra.MoveNext
										Loop 

										rsOra.close
										'##### Display page navigator #####
										sqlOra = "select count(*) as recordCount " _
										& " from ppo, ppoItem, aemployee , planreceive , planreceiveitem , psection , pdepartment , pdivision where " _
										& whereClause
										rsOra.open sqlOra, conn, adOpenStatic, adLockReadOnly, adCmdText
									
									'		recordCount = currentRecord - 1 
											recordCount = rsOra("recordCount")
										%>
                                        <tr> 
                                            <td colspan="13" height="2" align="center"><img src="images/spacer.gif" height="2"></td>
                                        </tr>
                                        <tr> 
                                            <td colspan="13" align="center"> 
                                                <%
											'Place Here
											displayPageNavigator absolutePage, recordCount, pageSize, request.serverVariables("SCRIPT_NAME") & "?ListDivision=" & request.queryString("ListDivision") & "&ListDepartment=" & request.queryString("ListDepartment") & "&ListSection=" & request.queryString("ListSection") & "&from_day=" & request.queryString("from_day") & "&from_month=" & request.queryString("from_month") & "&from_year=" & request.queryString("from_year") & "&to_day=" & request.queryString("to_day") & "&to_month=" & request.queryString("to_month") & "&to_year=" & request.queryString("to_year") & "&ab="
											%>
                                            </td>
                                        </tr>
                                        <%
									
											rsOra.close
											'##### END Display page navigator #####
									else
									%>
                                        <tr> 
                                            <td align="center" colspan="13"><br>
                                                <font><b>Data not found!</b></font><br>
                                                &nbsp;</td>
                                        </tr>
                                        <%
									End if
									%>
                                    </table>
                                </td>
                            </tr></Form>
                            <form name="c" aaction="">
                                <tr align="left"> 
                                    <td colspan="2" bgcolor="#FFD553" align="center"> 
                                        <br>
                                        <input type="button" value="Back" onClick="javascript: window.location.href = 'planrecrpt.asp'">&nbsp;
										<input type="button" name="btnExport" value="Export Excel" onClick="expExc_onclick()">	
                                        <br>
                                        &nbsp; </td>
                                </tr>
									
                                <tr align="left"> 
                                    <td width="98%" bgcolor="#FFD553" align="center">&nbsp;</td>
                                    <td width="2%" align="right" bgcolor="#FFD553" valign="bottom"><img src="../images/y04.jpg" width="12" height="12"></td>
                                </tr>
                            </form>
                        </table>
                    </td>
                </tr>
            </table>
            <font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#666666"></font></td>
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
</BODY>
</HTML>

