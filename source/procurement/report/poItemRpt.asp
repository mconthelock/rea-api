<% Option Explicit %>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file ="../inc/report.inc" -->
<!-- #include file ="../inc/reportData.inc" -->
<!-- #include file ="../inc/misc.inc" -->
<!-- #include virtual ="/sys/inc/pageUtil.inc" -->
<!-- #include virtual ="/sys/inc/oraUtil.inc" -->
<!-- #include virtual ="/sys/inc/search.inc" -->
<%
Dim rsOra, sqlOra
Dim color
Dim i, j, k, l, m, poNo
Dim fromClause, whereClause
Dim row
Dim pageSize, absolutePage, currentRecord, recordCount
Dim totalByPo, grandTotal, total, catType, poRunNo, poDate
Dim msg
Dim po, pr, issue, due, vName, vTel, vFax, rName, sec, inv, itemOrder, remain, prod, totalPORec, totalPOOrder
Dim  showProd, showPo, showPr, showVend, showSec
Dim isDate, vendor, exp, expGrp, desc

server.scriptTimeOut = 1800 '30 minutes

Color="#99CCFF"

set rsOra = createObject("adodb.recordset")

'Set page size
if request.queryString("ps") = "" or request.queryString("ps") < 0 then
	pageSize = 20
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
<TITLE>P/O ITEM REPORT</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css" type="text/css">
<style>
font {font-family: ms sans serif; font-size: 9pt}
</style></HEAD>
<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0"> <Form name="frmMain"> 
<%showPageLoading%>
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
                        <div align="center"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">P/O 
                          ITEM</font></font></b></div>
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr align="center"> 
                <td colspan="2" bgcolor="#FFD553" valign="top"> 
                  <table border=1 width="100%" align="Center" cellspacing="0" cellpadding="1" height="18">
                    <tr bgcolor="#CCCCFF"> 
                      <td height="20" align="center" width="1%" rowspan="2"><b><font color="#000000">NO</font></b></td>
                      <td height="19" align="center" rowspan="2" nowrap><b><font color="#000000">ISSUED 
                        DATE</font></b></td>
                      <td height="19" align="center" rowspan="2" nowrap><b><font color="#000000">DUE 
                        DATE</font></b></td>
                      <td height="9" align="center" width="1%"><b><font color="#000000">P/O#</font></b></td>
                      <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">REQUESTER</font></b></td>
                      <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">SEC</font></b></td>
                      <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">VENDOR/TEL,FAX</font></b></td>
                      <td height="20" width="100%" align="center" rowspan="2"><b><font color="#000000">DESCRIPTION</font></b></td>
                      <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">QTY 
                        ORDER</font></b></td>
                      <td height="20" width="1%" align="center" rowspan="2"><font color="#000000"><b>RECEIVED<br>
                        </b></font></td>
                      <td height="20" width="1%" align="center" rowspan="2"><font color="#000000"><b>REMAIN</b></font></td>
                      <td height="20" width="1%" align="center" rowspan="2"><b><font color="#000000">INV</font></b></td>
                    </tr>
                    <tr bgcolor="#CCCCFF"> 
                      <td height="10" align="center" width="1%"><b><font color="#000000">P/R#</font></b></td>
                    </tr>
                    <%
										fromClause = "ppr, ppo, ppoitem, pvender, aEmployee, pProduct, pSection, (select sPoNo, sProdCode, sRecBy, nQuality from pReceive, pReceiveItem where pReceiveItem.sRecNo = pReceive.sRecNo) a "

										'########## GENERATE WHERE CLAUSE ##########
										whereClause = "ppoitem.spono = ppo.spono and ppo.cStatus <> '9' " _
										& "and pvender.svendcode = ppo.svendcode " _
										& "and aemployee.sempno = ppr.srequester " _
										& "and psection.sseccode = aemployee.sseccode " _
										& "and a.spono(+) = ppoitem.spono and a.sprodcode(+) = ppoitem.sprodcode " _
										& "and ppr.sprno = ppo.srefno " _
										& "and pproduct.sprodcode = ppoitem.sprodcode and pproduct.sExpType = ppr.sExpType " _
										& "and ppo.cstatus <> '9' and (ppo.cstatus = '2' or ppo.cstatus = '3') "

										'Po number
										if request.queryString("txtPO") <> "" then
											whereClause = whereClause & " and ppoitem.sPoNo like '" & getSearchStr(request.queryString("txtPO")) & "'"
										end if
										
										'Issue date (from)
										if request.queryString("selPOSDate") <> "" and request.queryString("selPOSMonth") <> "" and request.queryString("selPOSYear") <> "" then
											whereClause = whereClause & " and dPoDate >= to_date('" & request.queryString("selPOSDate") & "/" & request.queryString("selPOSMonth") & "/" & request.queryString("selPOSYear") & "', 'DD/MM/YYYY')"
										end if

										'Issue date (to)
										if request.queryString("selPOEDate") <> "" and request.queryString("selPOEMonth") <> "" and request.queryString("selPOEYear") <> "" then
											whereClause = whereClause & " and dPoDate <= to_date('" & request.queryString("selPOEDate")  & "/" & request.queryString("selPOEMonth") & "/" & request.queryString("selPOEYear") & "', 'DD/MM/YYYY')"
										end if
										
										'P/R
										if request.queryString("txtPR") <> "" then
											whereClause = whereClause & " and ppo.sRefNo like '" & getSearchStr(request.queryString("txtPR")) & "'"
										end if

										'Vendor
										if request.queryString("txtVendName") <> "" then
											whereClause = whereClause & " and pVender.sEName like '" & getSearchStr(request.queryString("txtVendName")) & "'"
										end if

										'Requester
										if request.queryString("txtReq") <> "" then
											whereClause = whereClause & " and aEmployee.sEmpNo like '" & getSearchStr(request.queryString("txtReq")) & "'"
										end if

										'Section
										if 1 = 1 then
											whereClause = whereClause & " and srefno in (select sprno from ppr where srequester in (select sempno from aemployee where 1 = 1 "
											
											if request("selSection") <> "" then
												whereClause = whereClause & "and sseccode = '" & request("selSection") & "'"
											end if
											
											if request("selDep") <> "" then
												whereClause = whereClause & "and sdepcode = '" & request("selDep") & "'"
											end if
											
											if request("selDiv") <> "" then
												whereClause = whereClause & "and sdivcode = '" & request("selDiv") & "'"
											end if
											
											whereClause = whereClause & ")) "
										end if

										'Tax-invoice
										if request.queryString("txtInv") <> "" then
											whereClause = whereClause & " and a.sRecBy like '" & getSearchStr(request.queryString("txtInv")) & "'"
										end if

										'Product name
										if request.queryString("txtProdName") <> "" then
											whereClause = whereClause & " and sEProdName like '" & getSearchStr(request.queryString("txtProdName")) & "'"
										end if
										'########## END GENERATE WHERE CLAUSE ##########
										
										'########## GOTO START RECORD ##########
										sqlOra = "select " _
										& "ppo.spono, srefno, dpodate, to_char(nvl(ppoitem.nquality, 0)) as itemOrder, to_char(nvl(a.nQuality, 0)) as rec, pvender.sename, stelno, sfax " _
										& ", aEmployee.sEmpNo, aEmployee.sName, replace(sSec, 'SEC', '') as sSec, ppr.sDelivery, ppoitem.sProdCode, pProduct.sEProdName, a.sRecBy " _
										& "from " _
										& fromClause _
										& "where " _
										& whereClause _
										& " order by ppo.spono, ppo.sRefNo, sEProdName"
										'response.write sqlOra & "<br>"
										'response.end
										'response.write "<Time = " & now & "><br>"
										rsOra.open sqlOra, conn
										'response.end
										'response.write "<Time = " & now & "><br>"
										'########## END GOTO START RECORD ##########
										If Not rsOra.EOF then
											recordCount = 1
											
											do while not rsOra.EOF and recordCount  < absolutePage * pageSize - pageSize + 1
												if po <> rsOra("sPoNo") then
													po = rsOra("sPoNo")
													recordCount = recordCount + 1
													
													do while po = rsOra("sPoNo")
														rsOra.moveNext
														
														if rsOra.EOF then
															exit do
														end if
													loop
												end if
											loop

											currentRecord = absolutePage * pageSize - pageSize + 1
											row = currentRecord
											po = ""
											
											totalPORec = 0
											totalPOOrder = 0
											
											do while not rsOra.EOF and row < absolutePage * pageSize + 1
												inv = rsOra("sRecBy")
												
												if prod <> rsOra("sProdCode") then
													prod = rsOra("sProdCode")
													totalPOOrder = totalPOOrder + rsOra("itemOrder")
													showProd = true
												else
													showProd = false
												end if

												totalPORec = totalPORec + rsOra("rec")

												if po <> rsOra("sPoNo") then
													po = rsOra("sPoNo")
													showPo = true
													
													if color = "FFCCCC" then
														color = "99CCFF"
													else
														color = "FFCCCC"
													end if
												else
													showPo = false
												end if
												
												if pr <> rsOra("sRefNo") then
													showPr = true
													pr = rsOra("sRefNo")
												else
													showPr = false
												end if

												if vName <> rsOra("sEName") then
													showVend = true
													vName = rsOra("sEName")
												else
													showVend = false
												end if

												if sec <> rsOra("sSec") then
													showSec = true
													sec = rsOra("sSec")
												else
													showSec = false
												end if
										%>
                    <tr bgcolor="#<%=color%>" valign="top"> 
                      <td width="1%" height="2" align="center" rowspan="2"><font color="#000000"><%if showPo then response.write row else response.write "&nbsp;" end if%></font></td>
                      <td rowspan="2" height="1"><font color="#000000"> 
                        <%if showPo then response.write rsOra("dPODate") else response.write "&nbsp;" end if%>
                        </font></td>
                      <td rowspan="2" height="1"><font color="#000000"> 
                        <%if showPr then response.write rsOra("sDelivery") else response.write "&nbsp;" end if%>
                        </font></td>
                      <td width="1%" height="0"><font color="#000000"> 
                        <%if showPo then response.write rsOra("spono") else response.write "&nbsp;" end if%>
                        </font></td>
                      <td width="1%" height="20" rowspan="2"><font color="#000000"> 
                        <%if showPr then response.write "(" & rsOra("sEmpNo") & ") " & rsOra("sName") else response.write "&nbsp;" end if%>
                        </font></td>
                      <td width="1%" height="20" rowspan="2"><font color="#000000"> 
                        <%if showPr or showSec then response.write rsOra("sSec") else response.write "&nbsp;" end if%>
                        </font></td>
                      <td width="1%" height="20" rowspan="2"><font color="#000000"> 
                        <%
						if showVend then
							response.write rsOra("sEName")
							
							if not isNull(rsOra("sTelNo")) and trim(rsOra("sTelNo")) <> "" then
								response.write " Tel: " & rsOra("sTelNo")
							end if
	
							if not isNull(rsOra("sFax")) and trim(rsOra("sFax")) <> "" then
								response.write " Fax: " & rsOra("sFax")
							end if
						else
							response.write "&nbsp;"
						end if
						%>
                        </font></td>
                      <td width="100%" height="20" rowspan="2"><font color="#000000"> 
                        <%if showProd then response.write rsOra("sEProdName") else response.write "&nbsp;" end if%>
                        </font></td>
                      <td width="1%" height="20" align="right" rowspan="2"><font color="#000000"> 
                        <%if showProd then response.write rsOra("itemOrder") else response.write "&nbsp;" end if%>
                        </font></td>
                      <td width="1%" height="20" align="right" rowspan="2"><font color="#000000"><%=rsOra("rec")%></font></td>
                      <td width="1%" height="20" align="right" rowspan="2">
					  
					  <%
											'row = row + 1
											rsOra.MoveNext
											
										'########## SHOW REMAIN ORDER ##########
										if rsOra.EOF then
					%><font color="#000000"><%=totalPOOrder - totalPORec%></font><%
										elseif (po <> rsOra("sPoNo") or (po = rsOra("sPoNo") and prod <> rsOra("sProdCode"))) then
										%><font color="#000000"><%=totalPOOrder - totalPORec%></font><%
												totalPORec = 0
												totalPOOrder = 0
										else
											response.write "&nbsp;"
										end if
										
										if not rsOra.EOF then
											if row  >= absolutePage * pageSize and po = rsOra("sPoNo") then
													'row = row - 1
											elseif po <> rsOra("sPoNo") then
												row = row + 1
											end if
										end if
										'########## END SHOW REMAIN ORDER ##########
					  %>
					  
					  </td>
                      <td width="1%" height="20" rowspan="2" nowrap><font color="#000000"><%=inv%>&nbsp;</font></td>
                    </tr>
                    <tr valign="top" bgcolor="#<%=color%>"> 
                      <td width="1%" height="1"><font color="#000000"> 
                        <%if showPr or showPo then response.write pr else response.write "&nbsp;" end if%>
                        </font></td>
                    </tr>
                    <%
										Loop 
										rsOra.close
										%>
                    <%
										'response.write "<Time = " & now & "><br>"

										'##### Display page navigator #####
										sqlOra = "select count(distinct ppoitem.sPoNo) as recordCount from " & fromClause & " where " & whereClause
										'response.write sqlOra & "<br>"
										rsOra.open sqlOra, conn
										'response.write "<Time = " & now & "><br>"
									
									'		recordCount = currentRecord - 1 
											recordCount = rsOra("recordCount")
										%>
                    <tr> 
                      <td colspan="12" align="center"> 
                        <%
										'response.write "<Time = " & now & "><br>"
											'Place Here
											displayPageNavigator absolutePage, recordCount, pageSize, request.serverVariables("SCRIPT_NAME") & "?txtPO=" & server.urlEncode(request("txtPO")) & "&selPOSDate=" & server.urlEncode(request("selPOSDate")) & "&selPOSMonth=" & server.urlEncode(request("selPOSMonth")) & "&selPOSYear=" & server.urlEncode(request("selPOSYear")) & "&selPOEDate=" & server.urlEncode(request("selPOEMonth")) & "&selPOEYear=" & server.urlEncode(request("selPOEYear")) & "&txtReq=" & server.urlEncode(request("txtReq")) & "&txtPR=" & server.urlEncode(request("txtPR")) & "&txtVendName=" & server.urlEncode(request("txtVendName")) & "&selSection=" & server.urlEncode(request("selSection")) & "&txtInv=" & server.urlEncode(request("txtInv")) & "&txtProdName=" & server.urlEncode(request("txtProdName")) & "&ab="
										'response.write "<Time = " & now & "><br>"
											%>
                      </td>
                    </tr>
                    <%
									
											'rsOra.close
											'##### END Display page navigator #####
									else
									%>
                    <tr> 
                      <td align="center" colspan="12" bgcolor="#FFFFFE"><br>
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
                  <td colspan="2" bgcolor="#FFD553" align="center"> <br>
                    <input type="button" value="Back" onClick="javascript: window.location.href = 'poItemQryFrm.asp'">
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
<%hidePageLoading%>
</BODY>
</HTML>
 
<%
conn.close

set rsOra = nothing
set conn = nothing

server.scriptTimeOut = 90
%>