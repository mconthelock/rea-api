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
<!-- #include virtual ="/sys/inc/date.inc" -->

<%
Dim rsOra, sqlOra
Dim color
Dim i, j, k, l, m, poNo
Dim fromClause, whereClause, groupBy
Dim row
Dim pageSize, absolutePage, currentRecord, recordCount
Dim totalByPo, grandTotal, total, catType, poRunNo, poDate, prodCode, prodName, amonth, nQty, sdate, edate
Dim msg
Dim po, pr, issue, due, vName, vTel, vFax, rName, sec, inv, itemOrder, remain, prod, totalPORec, totalPOOrder
Dim  showProd, showPo, showPr, showVend, showSec
Dim isDate, vendor, exp, expGrp, desc
Dim totalRecord

server.scriptTimeOut = 1800 '30 minute

Color="#99CCFF"

set rsOra = createObject("adodb.recordset")

'Set page size
if request("ps") = "" or request("ps") < 0 then
	pageSize = 20
else
	pageSize = int(request("ps"))
end if

'Set absolute page
if request("ab") = "" or request("ab") < 0 then
	absolutePage = 1
else
	absolutePage = int(request("ab"))
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
                      <td height="20" align="center" width="1%" nowrap><b><font color="#000000">NO</font></b></td>
                      <td height="20" align="center" width="1%" nowrap><b><font color="#000000">ISSUED 
                        DATE</font></b></td>
                      <td height="19" align="center" nowrap width="1%"><b><font color="#000000">P/O 
                        NO</font></b></td>
                      <td height="20" width="1%" align="center" nowrap><b><font color="#000000">PRODUCT 
                        CODE</font></b></td>
                      <td height="20" width="100%" align="center" nowrap><b><font color="#000000">PRODUCT 
                        NAME</font></b></td>
                      <td height="20" width="1%" align="center"><b><font color="#000000">QTY</font></b></td>
                      <td height="20" width="1%" align="center" nowrap><b><font color="#000000">UNIT</font></b></td>
                      <td height="20" width="1%" align="center" nowrap><b>PRICE</b></td>
                      <td height="20" width="1%" align="center" nowrap><b>CUR</b></td>
                    </tr>
                    <%
										'########## GENERATE WHERE CLAUSE ##########
										'Issue date (from)
										if request("selPOSDate") <> "" and request("selPOSMonth") <> "" and request("selPOSYear") <> "" then
											sdate = request("selPOSDate")  & "/" & request("selPOSMonth") & "/" & request("selPOSYear")
											whereClause = whereClause & " and dPoDate >= to_date('" & request("selPOSDate") & "/" & request("selPOSMonth") & "/" & request("selPOSYear") & "', 'DD/MM/YYYY')"
										end if

										'Issue date (to)
										if request("selPOEDate") <> "" and request("selPOEMonth") <> "" and request("selPOEYear") <> "" then
											edate = request("selPOEDate")  & "/" & request("selPOEMonth") & "/" & request("selPOEYear")
											whereClause = whereClause & " and dPoDate <= to_date('" & request("selPOEDate")  & "/" & request("selPOEMonth") & "/" & request("selPOEYear") & "', 'DD/MM/YYYY')"
										end if
										
										'Product code
										if request("txtProdCode") <> "" then
											whereClause = whereClause & " and sProdCode like '" & getSearchStr(request("txtProdCode")) & "'"
										end if

										'Product name
										if request("txtProdName") <> "" then
											whereClause = whereClause & " and sProdName like '" & getSearchStr(request("txtProdName")) & "'"
										end if

										whereClause = "and ppo.cstatus <> '9' and (ppo.cstatus = '2' or ppo.cstatus = '3') " & whereClause
										whereClause = "ppoitem.spono = ppo.spono and ppo.cstatus <> '9' and ppoitem.cStatus = '1' " & whereClause
										'########## END GENERATE WHERE CLAUSE ##########

										'########## GOTO START RECORD ##########
										fromClause = "ppo, ppoitem "

										sqlOra = "select ppoitem.spono, sprodcode, sprodname " _
										& ", sunitcode, sCurCode, to_date(dpodate, 'DD/MM/YY') as dpodate, to_char(dpodate, 'MM/YY') as amonth, to_number(nquality) as nqty, to_number(nPrice) as nPrice " _
										& "from " & fromClause & " where " & whereClause _
										& " order by dpodate, ppoitem.spono, sprodcode, sprodname, nprice, scurcode "
										'response.write sqlOra
										'response.end
										rsOra.open sqlOra, conn
										'########## END GOTO START RECORD ##########
										
										If Not rsOra.EOF then
											recordCount = 1
											
											do while not rsOra.EOF and recordCount  < absolutePage * pageSize - pageSize + 1
												recordCount = recordCount + 1
												rsOra.moveNext
											loop
											
											if not rsOra.EOF then
												amonth = rsOra("amonth")
											end if

											currentRecord = absolutePage * pageSize - pageSize + 1
											row = currentRecord
											po = ""
											
											do while not rsOra.EOF and row < recordCount + pageSize
											%>
												<tr bgcolor="#<%=color%>" valign="top"> 
												  <td width="1%" height="2" align="center"><font color="#000000"> 
													<%=row%></font></td>
												  <td width="1%" height="2" align="center"><font color="#000000"> 
													<%
													if poDate <> rsOra("dPoDate") then
														poDate = rsOra("dPoDate")
														response.write toMedDate(rsOra("dPoDate"))
													else
														response.write "&nbsp;"
													end if
													%> </font></td>
												  <td height="1"><font color="#000000"> <%
												  if po <> rsOra("sPoNo") then
													po = rsOra("sPoNo")
													response.write rsOra("sPoNo")
												  else
													response.write "&nbsp"
												  end if
												  %> 
													</font></td>
												  
                      <td width="1%" height="20" nowrap><font color="#000000"> 
                        <%
												  if prodCode <> rsOra("sProdCode") then
													prodCode = rsOra("sProdCode")
													response.write rsOra("sProdCode")
												  else
													response.write "&nbsp"
												  end if
												  %>
                        </font></td>
												  
                      <td width="1%" height="20" nowrap><font color="#000000"> 
                        <%
												  if prodName <> rsOra("sProdName") then
													prodName = rsOra("sProdName")
													response.write rsOra("sProdName")
												  else
													response.write "&nbsp"
												  end if
												  %>
                        </font></td>
												  <td width="1%" height="20" align="right"><font color="#000000"> 
													<%=formatNumber(rsOra("nQty") & "", 4)%></font></td>
												  <td width="1%" height="20"><font color="#000000"> 
													<%=rsOra("sUnitCode")%> </font></td>
												  <td width="1%" height="20" align="right"><font color="#000000"><%=formatNumber(rsOra("nPrice"), 4)%></font></td>
												  <td width="1%" height="20"><font color="#000000"><%=rsOra("sCurCode")%></font></td>
												</tr>
												<%
													row = row + 1
												rsOra.moveNext
												
												if rsOra.EOF then
													showRollup amonth, sdate, edate
												else
													if amonth <> rsOra("amonth") then
														showRollup amonth, sdate, edate
													end if

													amonth = rsOra("amonth")
												end if
											Loop 
											
											rsOra.close

										'response.write "<Time = " & now & "><br>"

										'##### Display page navigator #####
										sqlOra = "select count(*) as recordCount from " & fromClause & " where " & whereClause
										'response.write sqlOra & "<br>"
										rsOra.open sqlOra, conn
										'response.write "<Time = " & now & "><br>"
									
									'		recordCount = currentRecord - 1 
											recordCount = rsOra("recordCount")
										%>
                    <tr> 
                      <td colspan="9" align="center"> 
                        <%
										'response.write "<Time = " & now & "><br>"
											'Place Here
											displayPageNavigator absolutePage, recordCount, pageSize, request.serverVariables("SCRIPT_NAME") & "?selPOSDate=" & server.urlEncode(request("selPOSDate")) & "&selPOSMonth=" & server.urlEncode(request("selPOSMonth")) & "&selPOSYear=" & server.urlEncode(request("selPOSYear")) & "&selPOEMonth=" & server.urlEncode(request("selPOEMonth")) & "&selPOEDate=" & server.urlEncode(request("selPOEDate")) & "&selPOEYear=" & server.urlEncode(request("selPOEYear")) & "&txtProdCode=" & server.urlEncode(request("txtProdCode")) & "&txtProdName=" & server.urlEncode(request("txtProdName")) & "&ab="
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
                      <td align="center" colspan="9" bgcolor="#FFFFFE"><br>
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
                    <input type="button" value="Back" onClick="javascript: window.location.href = 'sum_prod.asp'">
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
'response.write whereClause
conn.close

set rsOra = nothing
set conn = nothing

server.scriptTimeOut = 90

function showRollup(amonth, sdate, edate)
	dim rsOra, sqlOra
	
	set rsOra = createObject("adodb.recordset")
	sqlOra = "select " _
	& "to_char(dpodate, 'MM/YY') as amonth, scurcode, sunitcode, sum(nquality) as nqty, sum(nPrice) as nPrice " _
	& ", (grouping(to_char(dpodate, 'MM/YY')) || grouping(scurcode) || grouping(sunitcode)) as totalflag " _
	& "from ppo, ppoitem " _
	& "where ppoitem.spono = ppo.spono and ppoitem.cStatus = '1' and dPoDate >= to_date('" & sdate & "', 'DD/MM/YYYY') " _
	& "and dPoDate <= to_date('" & edate & "', 'DD/MM/YYYY') " _
	& "and to_char(dpodate, 'MM/YY') = '" & amonth & "' " _
	& "group by rollup(to_char(dpodate, 'MM/YY'), scurcode, sunitcode) " _
	& "order by sunitcode, scurcode "

	sqlOra = "select " _
	& "to_char(dpodate, 'MM/YY') as amonth, scurcode, sunitcode, sum(nquality) as nqty, sum(nPrice) as nPrice " _
	& ", (grouping(to_char(dpodate, 'MM/YY')) || grouping(scurcode) || grouping(sunitcode)) as totalflag " _
	& "from ppo, ppoitem " _
	& "where " & whereClause _
	& "and to_char(dpodate, 'MM/YY') = '" & amonth & "' " _
	& "group by rollup(to_char(dpodate, 'MM/YY'), scurcode, sunitcode) " _
	& "order by sunitcode, scurcode "
	'response.write sqlOra
	rsOra.open sqlOra, conn

	do while not rsOra.EOF
		if rsOra("totalflag") = "000" then
			response.write "<tr><td colspan=""5"" width=""100%"">&nbsp;</td><td align=""right"" width=""1%""><font color=""000000"">" & formatNumber(rsOra("nQty") & "", 4) & "</font></td><td width=""1%""><font color=""000000"">" & rsOra("sunitcode") & "</font></td><td align=""right"" width=""1%""><font color=""00000"">" & formatNumber(rsOra("nPrice") & "", 4) & "</font></td><td width=""1%""><font color=""00000"">"  & rsOra("scurcode") & "</font></td></tr>"
		end if
		rsOra.moveNext
	loop
	
	rsOra.close
	set rsOra = nothing
end function
%>