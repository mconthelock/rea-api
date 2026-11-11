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
Dim whereClause, groupClause
Dim row
Dim pageSize, absolutePage, currentRecord, recordCount
Dim totalByPo, grandTotal, total, catType, poRunNo, poDate
Dim cur, curTotal, grp, grpTotal, unit, unitTotal
Dim msg
Dim vend, vendName, po, rcvDate

server.scriptTimeOut = 1800 '30 minute

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
%><HTML>
<HEAD> 
<TITLE>P/O ISSUED REPORT</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
<style>
font {font-family: ms sans serif; font-size: 9pt}
</style>
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0"> <Form name="frmMain"> 
<table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
        <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
        <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    </tr>
    <tr> 
        <td width="1%">&nbsp;</td>
        <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><font color="#BF00BF">Report 
            Manager</font></b></font></td>
    </tr>
    <tr> 
        <td width="1%"><img src="../images/spacer.gif" width="1" height="1"></td>
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
                                                <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">MITSUBISHI 
                                                    ELEVATOR ASIA CO.,LTD.<br>
                                                    PURCHASE ORDER OUTSTANDING 
                                                    SUMMARY REPORT</font></font></b></div>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr> 
                                <td colspan="2" bgcolor="silver"> <font color="blue"><b>ISSUE 
                                    DATE: </b><%=request.queryString("from_day") & "/" & request.queryString("from_month") & "/" & request.queryString("from_year") & " - " & request.queryString("to_day") & "/" & request.queryString("to_month") & "/" & request.queryString("to_year")%></font></td>
                            </tr>
                            <tr align="center"> 
                                <td colspan="2" bgcolor="#FFD553" valign="top"> 
                                    <table border=0 width="100%" align="Center" cellspacing="1" cellpadding="2" height="18">
                                        <tr bgcolor="#CCCCFF"> 
                                            <td height="20" align="center" width="1%" rowspan="2"><b><font color="#000000">NO</font></b></td>
                                            <td height="20" align="center" width="1%" rowspan="2"><b><font color="#000000">VEND</font></b></td>
                                            <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">RCV<br>
                                                DATE</font></b></td>
                                            <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">P/O#</font></b></td>
                                            <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">CUR</font></b></td>
                                            <td height="9" colspan="4" align="center"><b><font color="#000000">DELIVERY 
                                                ON</font></b></td>
                                            <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">P/O 
                                                DELIVERY ON DATE</font></b></td>
                                            <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">REMARK</font></b></td>
                                        </tr>
                                        <tr bgcolor="#CCCCFF"> 
                                            <td height="10" width="0%" align="center"><b><font color="#000000">THIS 
                                                MONTH AND BEFORE</font></b></td>
                                            <td height="10" width="1%" align="center"><b><font color="#000000">30 
                                                DAYS </font></b></td>
                                            <td height="10" width="1%" align="center"><b><font color="#000000">60 
                                                DAYS </font></b></td>
                                            <td height="10" width="1%" align="center"><b><font color="#000000">90 
                                                DAYS </font></b></td>
                                        </tr>
                                        <%
										'########## GENERATE WHERE CLAUSE ##########
										whereClause = "1 = 1 " _
										& "and ppo.cstatus <> '9' and (ppo.cstatus = '2' or ppo.cstatus = '3') "
										
										'Po number
										if request.queryString("po") <> "" then
											whereClause = whereClause & " and ppoitem.sPoNo = '" & request.queryString("po") & "'"
										end if
										
										'Vendor
										if request.queryString("vend") <> "" then
											whereClause = whereClause & " and ppo.sVendCode = '" & request.queryString("vend") & "'"
										end if
										
										'Issue date (from)
										if request.queryString("from_day") <> "" then
											whereClause = whereClause & " and dPoDate >= to_date('" & request.queryString("from_day") & "/" & request.queryString("from_month") & "/" & request.queryString("from_year") & "', 'DD/MM/YYYY')"
										end if

										'Issue date (to)
										if request.queryString("to_day") <> "" then
											whereClause = whereClause & " and dPoDate <= to_date('" & request.queryString("to_day")  & "/" & request.queryString("to_month") & "/" & request.queryString("to_year") & "', 'DD/MM/YYYY')"
										end if
										'########## END GENERATE WHERE CLAUSE ##########
										
										groupClause = "ppo.sPoNo, ppo.dPoDate, sCurrency, ppo.sVendCode " _
										& ", pVender.sEName, sDelivery "
										
										'########## GOTO START RECORD ##########
										sqlOra = "select ppo.sPoNo, ppo.dPoDate, ppo.sVendCode, sDelivery " _
										& ", sum(ppoItem.nPrice * ppoItem.nQuality) as nPrice, sCurrency, pVender.sEName " _
										& "from ppo, ppoItem, pUnit, pVender " _
										& "where " _
										& whereClause _
										& "and ppoitem.spono = ppo.spono " _
										& "and pVender.sVendCode = ppo.sVendCode and pUnit.sUnitCode = ppoitem.sUnitCode " _
										& "and ppo.sRefNo not in (select distinct sPrNo from pReceive) " _
										& " group by " & groupClause _
										& " order by dPoDate, ppo.sPoNo, ppo.sVendCode"
										'response.write sqlOra & "<br>"
										rsOra.open sqlOra, conn
									
										recordCount = 1
										
										do while not rsOra.EOF and recordCount  < absolutePage * pageSize - pageSize + 1
											recordCount = recordCount + 1
											rsOra.moveNext
										loop
									
										currentRecord = absolutePage * pageSize - pageSize + 1
										row = currentRecord
										
										'########## END GOTO START RECORD ##########
										grandTotal = 0
										
										If Not rsOra.EOF then
											i = 1
											poRunNo = 1
											poNo = rsOra("sPoNo")
											j = 0
											l = 0
											m = 0
											
											'cur = lCase(rsOra("sCurrency"))
											
											redim grp(0)
											redim grpTotal(0)
											redim cur(0)
											redim curTotal(0)
											redim unit(0)
											redim unitTotal(0)

											do while not rsOra.EOF and row < absolutePage * pageSize + 1
												'########## SUM CUR GROUP ##########
												if inStr(join(cur, ","), rsOra("sCurrency")) = 0 then
													l = l + 1
													redim preserve cur(l)
													redim preserve curTotal(l)
													cur(l) = rsOra("sCurrency")
													curTotal(l) = rsOra("nPrice")
												else
													for k = 1 to uBound(cur)
														if cur(k) = rsOra("sCurrency") then
															curTotal(k) = cLng(curTotal(k)) + cLng(rsOra("nPrice"))
														end if
													next
												end if
												'########## END SUM CUR GROUP ##########

										%>
                                        <tr bgcolor="<%=Color%>" valign="top"> 
                                            <td width="1%" height="2" align="center"><font color="#000000"><%=row%></font></td>
                                            <td width="1%" height="2"> <font color="#000000"> 
                                                <%vend = getValue(vend, rsOra("sVendCode"))%>
                                                <br>
                                                <%vendName = getValue(vendName, rsOra("sEName"))%>
                                                </font></td>
                                            <td width="1%" height="20" align="center"> 
                                                <font color="#000000"> 
                                                <%rcvDate = getValue(rcvDate, rsOra("dPoDate"))%>
                                                </font></td>
                                            <td width="1%" height="20"> <font color="#000000"> 
                                                <%po = getValue(po, rsOra("sPoNo"))%>
                                                </font></td>
                                            <td width="1%" height="20" align="center"><font color="#000000"><%=rsOra("sCurrency")%></font></td>
                                            <td width="0%" height="20" align="right"><font color="#000000"></font></td>
                                            <td width="0%" height="20" align="right"><font color="#000000"></font></td>
                                            <td width="0%" height="20" align="right"><font color="#000000"></font></td>
                                            <td width="1%" height="20" align="right"><font color="#000000"></font></td>
                                            <td width="1%" height="20" align="center"><font color="#000000"><%=rsOra("sDelivery")%></font></td>
                                            <td width="1%" height="20"><font color="#000000"><%=rsOra("sDelivery")%></font></td>
                                        </tr>
                                        <%
											if Color="#FFCCCC" then
												Color="#99CCFF"
											else
												Color="#FFCCCC"
											end if
											
											Color = "#99CCFF"
											
											i = i + 1
											row = row + 1
											grandTotal = grandTotal + cLng(rsOra("nPrice"))
											rsOra.MoveNext
										Loop

										rsOra.close
										
										'grandTotal = grandTotal + totalByPo
										%>
                                        <!--########## TOTAL BY CURRENCY ##########-->
                                        <%showCurSummary%>
                                        <!--########## END TOTAL BY CURRENCY ##########-->
                                        <!--########## GRAND TOTAL ##########-->
                                        <tr> 
                                            <td colspan="16"> 
                                                <hr size="1">
                                            </td>
                                        </tr>
                                        <tr> 
                                            <td colspan="9">GRAND TOTAL</td>
                                            <td align="center"> 
                                                <%'=rsOra("sUnitDesc")%>
                                            </td>
                                            <td colspan="2" align="right"><%=formatNumber(grandTotal, 4)%></td>
                                        </tr>
                                        <!--########## END GRAND TOTAL ##########-->
                                        <%

										'##### Display page navigator #####
										sqlOra = "select count(*) as recordCount from ppo where " & whereClause _
										& "and ppo.sRefNo not in (select distinct sPrNo from pReceive) " _
										'response.write sqlOra & "<br>"
										rsOra.open sqlOra, conn, adOpenStatic, adLockReadOnly, adCmdText
									
									'		recordCount = currentRecord - 1 
											recordCount = rsOra("recordCount")
										%>
                                        <tr> 
                                            <td colspan="16" height="2" align="center"><img src="images/spacer.gif" height="2"></td>
                                        </tr>
                                        <tr> 
                                            <td colspan="16" align="center"> 
                                                <%
											'Place Here
												displayPageNavigator absolutePage, recordCount, pageSize, request.serverVariables("SCRIPT_NAME") & "?po=" & request.queryString("po") & "&vend=" & request.queryString("vend") & "&exp=" & request.queryString("exp") & "&from_day=" & request.queryString("from_day") & "&from_month=" & request.queryString("from_month") & "&from_year=" & request.queryString("from_year") & "&to_day=" & request.queryString("to_day") & "&to_month=" & request.queryString("to_month") & "&to_year=" & request.queryString("to_year") & "&ab="
											%>
                                            </td>
                                        </tr>
                                        <%
									
											rsOra.close
											'##### END Display page navigator #####
									else
									%>
                                        <tr> 
                                            <td align="center" colspan="16"><br>
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
                                        <input type="button" value="Back" onClick="javascript: window.location.href = 'poOutSumRptQry.asp'">
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
</table></BODY>
</HTML> 
<%
conn.close

set rsOra = nothing
set conn = nothing

server.scriptTimeOut = 90

function showCurSummary()
	Dim rsOra, sqlOra
	Dim msg
	Dim i
	
	msg = "TOTAL BY CURRENCY"
	%>
<tr> 
    <td colspan="13" align="right"> 
        <hr size="1">
    </td>
</tr>
<%
	for i = 1 to uBound(cur)
	%>
<tr> 
    <td colspan="9" align="left"><%=msg%></td>
    <td align="left"><%=cur(i)%></td>
    <td colspan="2" align="right"><%=formatNumber(curTotal(i), 4)%></td>
</tr>
<%
		msg = "&nbsp;"
	next
	
	exit function
	set rsOra = createObject("adodb.recordset")
%>
<tr> 
    <td colspan="13"> 
        <hr size="1">
    </td>
</tr>
<%
	sqlOra = "select sCurrency, sum(ppoItem.nQuality * ppoItem.nPrice) as total " _
	& "from ppo, ppoItem, pProduct, pUnit " _
	& "where " & whereClause _
	& " group by sCurrency" _
	& " order by sCurrency"
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, conn

	msg = "TOTAL BY CURRENCY"
		
	do while not rsOra.EOF
%>
<tr align="center"> 
    <td colspan="6" align="right"><%=msg%></td>
    <td align="left"><%=rsOra("sCurrency")%></td>
    <td colspan="4" align="right"><%=formatNumber(rsOra("total"), 4)%></td>
</tr>
<%
		msg = "&nbsp;"
		rsOra.moveNext
	loop
	
	rsOra.close
end function
%>
