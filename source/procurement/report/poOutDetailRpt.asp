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
Dim vend, vendName, po, exp, expGrp, desc, rcvDate

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
%>

<HTML>
<HEAD>
<TITLE>P/O ISSUED REPORT</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
		
<Form name="frmMain">
  <table width="98%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="left" valign="top" width="1%" height="2"><img src="../images/ccb.gif" width="10" height="10"></td>
      <td width="99%" height="2" align="center" valign="top"><img src="../images/spacer.gif" width="8" height="8"></td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="99%" height="30" valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="+1" color="#666666"><b><font color="#BF00BF">Report Manager</font></b></font></td>
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
                                                    ITEM LIST REPORT</font></font></b></div>
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
                                            <td height="20" align="center" width="1%"><b><font color="#000000">NO</font></b></td>
                                            <td height="20" align="center" width="1%"><b>P/O#</b></td>
                                            <td height="20" width="1%" align="center"><b>RCV<br>
                                                DATE</b></td>
                                            <td height="20" width="1%" align="center"><b>VEND</b></td>
                                            <td height="20" width="1%" align="center"><b>EXP. 
                                                TYPE</b></td>
                                            <td height="20" width="1%" align="center"><b>EXP. 
                                                GROUP </b></td>
                                            <td height="20" width="1%" align="center"><b> 
                                                #</b></td>
                                            <td height="20" align="center"><b>DESCRIPTION</b></td>
                                            <td height="20" width="1%" align="center"><b>UNIT</b></td>
                                            <td height="20" width="1%" align="center"><b>QTY</b></td>
                                            <td height="20" width="1%" align="center"><b>CUR</b></td>
                                            <td height="20" width="1%" align="center"><b>@</b></td>
                                            <td height="20" width="1%" align="center"><b>TOTAL</b></td>
                                        </tr>
                                        <%
										'########## GENERATE WHERE CLAUSE ##########
										whereClause = "ppoitem.spono = ppo.spono and ppoitem.sprodcode = pproduct.sprodcode " _
										& "and ppo.svendcode = pproduct.svendcode and ppoitem.sunitcode = punit.sunitcode " _
										& "and ppo.sVendCode = pVender.sVendCode " _
										& "and ppo.cstatus <> '9' and (ppo.cstatus = '2' or ppo.cstatus = '3') "
										
										'Po number
										if request.queryString("po") <> "" then
											whereClause = whereClause & " and ppoitem.sPoNo = '" & request.queryString("po") & "'"
										end if
										
										'Expense type
										if request.queryString("exp") <> "" then
											whereClause = whereClause & " and pProduct.sExpType = '" & request.queryString("exp") & "'"
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
										
										'########## GOTO START RECORD ##########
										sqlOra = "select " _
										& "ppoItem.sPoNo, ppo.dPoDate, ppo.sVendCode, sProdName, pProduct.sExpType, " _
										& "pProduct.sCatType, sUnitDesc, to_char(ppoItem.nQuality) as nQuality, " _
										&"to_char(ppoItem.nPrice) as nPrice, sCurrency, pVender.sEName " _
										& "from ppo, ppoItem, pProduct, pUnit, pVender " _
										& "where " & whereClause _
										& " order by dPoDate, sPoNo, ppo.sVendCode, sProdName, sUnitDesc"
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
									
										If Not rsOra.EOF then
											i = 0
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
												'########## SUM EXPENSE GROUP ##########
												if inStr(join(grp, ","), rsOra("sCatType")) = 0 then
													j = j + 1
													redim preserve grp(j)
													redim preserve grpTotal(j)
													grp(j) = rsOra("sCatType")
													grpTotal(j) = rsOra("nQuality") * rsOra("nPrice")
												else
													for k = 1 to uBound(grp)
														if grp(k) = rsOra("sCatType") then
															grpTotal(k) = grpTotal(k) + rsOra("nQuality") * rsOra("nPrice")
														end if
													next
												end if
												'########## END SUM EXPENSE GROUP ##########

												'########## SUM CUR GROUP ##########
												if inStr(join(cur, ","), rsOra("sCurrency")) = 0 then
													l = l + 1
													redim preserve cur(l)
													redim preserve curTotal(l)
													cur(l) = rsOra("sCurrency")
													curTotal(l) = rsOra("nQuality") * rsOra("nPrice")
												else
													for k = 1 to uBound(cur)
														if cur(k) = rsOra("sCurrency") then
															curTotal(k) = curTotal(k) + rsOra("nQuality") * rsOra("nPrice")
														end if
													next
												end if
												'########## END SUM CUR GROUP ##########

												'########## TOTAL BY P/O ##########
												if rsOra("sPoNo") <> poNo then
												%>
                                        <!--########## TOTAL BY UNIT ##########-->
                                        <%
												msg = "TOTAL BY UNIT"
												
												for i = 0 to uBound(unit)
												%>
                                        <tr> 
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left"><%=msg%></td>
                                            <td align="center"><%=unit(i)%></td>
                                            <td align="left">&nbsp;</td>
                                            <td colspan="4" align="right"><%=formatNumber(unitTotal(i), 4)%></td>
                                        </tr>
                                        <%
													msg = "&nbsp;"
												next
												
												redim unit(0)
												redim unitTotal(0)
												%>
                                        <!--########## END TOTAL BY UNIT ##########-->
                                        <tr> 
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>TOTAL BY P/O</td>
                                            <td>&nbsp;</td>
                                            <td align="center">&nbsp;</td>
                                            <td colspan="4" align="right"><%=formatNumber(totalByPo, 4)%></td>
                                        </tr>
                                        <%
													grandTotal = grandTotal + totalByPo
													totalByPo = rsOra("nQuality") * rsOra("nPrice")
													poNo = rsOra("sPoNo")
													i = 1
													m = 0
												else
													totalByPo = totalByPo + rsOra("nQuality") * rsOra("nPrice")
													i = i + 1
												end if
												'########## END TOTAL BY P/O ##########
												
												'########## SUM UNIT ##########
												if inStr(join(unit, ","), rsOra("sUnitDesc")) = 0 then
													redim preserve unit(m)
													redim preserve unitTotal(m)
													unit(m) = rsOra("sUnitDesc")
													unitTotal(m) = rsOra("nQuality") * rsOra("nPrice")
													m = m + 1
												else
													for k = 0 to uBound(unit)
														if unit(k) = rsOra("sUnitDesc") then
															unitTotal(k) = unitTotal(k) + rsOra("nQuality") * rsOra("nPrice")
														end if
													next
												end if
												'########## END SUM UNIT ##########

										%>
                                        <tr bgcolor="#FFCCCC" valign="top"> 
                                            <td width="1%" bgcolor="<%=Color%>" height="2" align="center"><%=row%></td>
                                            <td width="1%" bgcolor="<%=Color%>" height="2"> 
                                                <%po = getValue(po, rsOra("sPoNo"))%>
                                            </td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="center"> 
                                                <%rcvDate = getValue(rcvDate, rsOra("dPoDate"))%>
                                            </td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>"> 
                                                <%vend = getValue(vend, rsOra("sVendCode"))%>
                                                <br>
                                                <%vendName = getValue(vendName, rsOra("sEName"))%>
                                            </td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="center" > 
                                                <%exp = getValue(exp,rsOra("sExpType"))%>
                                            </td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="center"> 
                                                <%expGrp = getValue(expGrp, rsOra("sCatType"))%>
                                            </td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="center"><%=i%></td>
                                            <td height="20" bgcolor="<%=Color%>"> 
                                                <%desc = getValue(desc, rsOra("sProdName"))%>
                                            </td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="center"><%=rsOra("sUnitDesc")%></td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="right"><%=rsOra("nQuality")%></td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="center"><%=rsOra("sCurrency")%></td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="right"><%=FormatNumber(rsOra("nPrice"), 4)%></td>
                                            <td width="1%" height="20" bgcolor="<%=Color%>" align="right"><%=FormatNumber(rsOra("nQuality") * rsOra("nPrice"), 4)%></td>
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
										
										grandTotal = grandTotal + totalByPo
										%>
                                        <!--########## TOTAL BY UNIT ##########-->
                                        <%for i = 0 to uBound(unit)%>
                                        <tr> 
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">TOTAL BY UNIT</td>
                                            <td align="center"><%=unit(i)%></td>
                                            <td align="left">&nbsp;</td>
                                            <td colspan="4" align="right"><%=formatNumber(unitTotal(i), 4)%></td>
                                        </tr>
                                        <%next%>
                                        <!--########## END TOTAL BY UNIT ##########-->
                                        <!--########## TOTAL BY P/O ##########-->
                                        <tr> 
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">TOTAL BY P/O</td>
                                            <td align="left">&nbsp;</td>
                                            <td align="left">&nbsp;</td>
                                            <td colspan="4" align="right"><%=formatNumber(totalByPo, 4)%></td>
                                        </tr>
                                        <!--########## END TOTAL BY P/O ##########-->
                                        <!--########## TOTAL BY CURRENCY ##########-->
                                        <%showCurSummary%>
                                        <!--########## END TOTAL BY CURRENCY ##########-->
                                        <!--########## TOTAL BY EXPENSE TYPE ##########-->
                                        <%showExpSummary%>
                                        <!--########## END TOTAL BY EXPENSE TYPE ##########-->
                                        <!--########## GRAND TOTAL ##########-->
                                        <tr> 
                                            <td colspan="13"> 
                                                <hr size="1">
                                            </td>
                                        </tr>
                                        <tr> 
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>&nbsp;</td>
                                            <td>GRAND TOTAL</td>
                                            <td> 
                                                <%'=rsOra("sUnitDesc")%>
                                            </td>
                                            <td>&nbsp;</td>
                                            <td colspan="4" align="right"><%=formatNumber(grandTotal, 4)%></td>
                                        </tr>
                                        <!--########## END GRAND TOTAL ##########-->
                                        <%

										'##### Display page navigator #####
										sqlOra = "select count(*) as recordCount from ppo, ppoitem, pproduct, punit, pVender where " & whereClause
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
                                        <input type="button" value="Back" onClick="javascript: window.location.href = 'poOutDetailRptQry.asp'">
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
		<td colspan="13" align="right"><hr size="1"></td>
	</tr>

	<%
	for i = 1 to uBound(cur)
	%>

	<tr>
		<td colspan="7" align="left">&nbsp;</td>
		<td align="left"><%=msg%></td>
		<td align="center"><%=cur(i)%></td>
		<td colspan="4" align="right"><%=formatNumber(curTotal(i), 4)%></td>
	</tr>

	<%
		msg = "&nbsp;"
	next
	
	exit function
	set rsOra = createObject("adodb.recordset")
%>

<tr>
	<td colspan="13"><hr size="1"></td>
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
	<td colspan="8" align="right">&nbsp;</td>
	<td  align="right"><%=msg%></td>
	<td align="left"><%=rsOra("sCurrency")%></td>
	<td colspan="4" align="right"><%=formatNumber(rsOra("total"), 4)%></td>
</tr>

<%
		msg = "&nbsp;"
		rsOra.moveNext
	loop
	
	rsOra.close
end function

function showExpSummary()
	Dim rsOra, sqlOra
	Dim catType, msg
	Dim i

	msg = "TOTAL BY EXPENSE TYPE GROUP"
	%>

	<tr>
		<td colspan="13" align="right"><hr size="1"></td>
	</tr>

	<%
	for i = 1 to uBound(grp)
	%>
	<tr>
		<td colspan="7" align="left">&nbsp;</td>
		<td align="left"><%=msg%></td>
		<td align="left"><%=grp(i)%></td>
		<td colspan="4" align="right"><%=formatNumber(grpTotal(i), 4)%></td>
	</tr>
	<%
		msg = "&nbsp;"
	next
	
	exit function
	
	set rsOra = createObject("adodb.recordset")
%>
<tr>
	<td colspan="13"><hr size="1"></td>
</tr>

<%
	sqlOra = "select pproduct.sCatType, sum(ppoItem.nQuality * ppoItem.nPrice) as total " _
	& "from ppo, ppoItem, pProduct, pUnit " _
	& "where " & whereClause _
	& " group by sCatType " _
	& " order by sCatType"
	'response.write sqlOra & "<br>"
	rsOra.open sqlOra, conn
	
	catType = rsOra("sCatType")
	total = 0
	
	do while not rsOra.EOF
%>
<tr align="center"> 
	<td colspan="8" align="right">&nbsp;</td>
	<td align="right">TOTAL BY GROUP</td>
	<td align="left"><%=rsOra("sCatType")%></td>
	<td colspan="4" align="right"><%=formatNumber(rsOra("total"), 4)%></td>
</tr>

<%
		rsOra.moveNext
	loop
	
	rsOra.close
end function
%>