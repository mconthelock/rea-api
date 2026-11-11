<%
Option Explicit
Dim conn, rs, sql, i
Dim ppoItems, dpoDate, dpoMonth, sProdCode, sProdName, sUnitCode
Dim qryDpoDate, qrySProdCode, qrySProdName, qrySort, extWhere
Dim mProduct, mQty, totalProduct, totalQty
Dim pageSize, pageCount, pageNo

mProduct = 0
mQty = 0
totalProduct = 0
totalQty = 0
i = 0
%>

<!-- #include file="../adovbs.inc" -->
<!-- #include file="../calender.inc" -->
<!-- #include file="../select.inc" -->

<html>
<head>
<style>
font
{
	font-family: ms sans serif;
	font-size: 10pt
}

a
{
	text-decoration: none
}
</style>
<title>PPOITEM Report</title></head>
<body>
<%
	'Set query
	qryDpoDate = session("qryDpoDate")
	
	if (int(request("selSDate")) >= 1 and int(request("selSDate")) <= 31) and (int(request("selFDate")) >= 1 and int(request("selFDate")) <= 31) then
		if (int(request("selSMonth")) >= 1 and int(request("selSMonth")) <= 12) and (int(request("selFMonth")) >= 1 and int(request("selFMonth")) <= 12) then
			if (int(request("selSYear")) >= 1992 and int(request("selSYear")) <= year(formatDateTime(date, vbShortDate))) and (int(request("selFYear")) >= 1992 and int(request("selFYear")) <= year(formatDateTime(date, vbShortDate))) then
				qryDpoDate = "and dpoDate >= '" & LCase(request("selSDate") & "-" & m(request("selSMonth"), 1) & "-" & right(request("selSYear"), 2)) & "' and  dpoDate <= '" & LCase(request("selFDate") & "-" & m(request("selFMonth"), 1) & "-" & right(request("selFYear"), 2)) & "'"
			end if
		end if
	end if
	
	qryDpoDate = session("qryDpoDate")
	qrySProdCode = session("qrySProdCode")
	
	if request("txtProdCode") <> "" then
		qrySProdCode = "and sProdCode like '%" & request("txtProdCode") & "%'"
'	else
'		qrySProdCode = ""
	end if
	
	session("qrySProdCode") = qrySProdCode
	qrySProdName = session("qrySProdName")
	
	if request("txtProdName") <> "" then
		qrySProdName = "and sProdName like '%" & request("txtProdName") & "%'"
	else
		qrySProdName = ""
	end if
	
	session("qrySProdName") = qrySProdName
	
	'Set sort by column
	qrySort = session("qrySort")
	
	if request.queryString("sort") <> "" then
		qrySort = request.queryString("sort")
	elseif request("selSort") <> "" then
		qrySort = request("selSort")
	else
		qrySort = "dpoDate, sProdCode"
	end if
	
	session("qrySort") = qrySort
'	response.write "sort " & session("qrySort") & "prodcode " & session("qrySProdCode") & "name " & session("qrySProdName")

'If no query then don't display any data
if qryDpoDate <> "" or qrySProdCode <> "" or qrySProdName <> "" then
%>
<table cellpadding="0" cellspacing="0" border="0" align="center" width="100%" bgcolor="#FFCC66">
	<tr align="center"> 
		<td align="left" valign="top"><img src="images/y01.jpg" width="12" height="12"></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
		<td align="right" valign="top"><img src="images/y02.jpg" width="12" height="12"></td>
	</tr>
	<tr align="center"> 
		<td></td>
		<td bgcolor="#CC6633" align="left" valign="top"><img src="images/hCorner1.jpg" width="12" height="12"></td>
		<td bgcolor="#CC6633" colspan="4" valign="middle"><b><font color="#FFFFFF">PPOITEM 
			Report</font></b></td>
		<td bgcolor="#CC6633" align="right" valign="top"><img src="images/hCorner2.jpg" width="12" height="12"></td>
		<td></td>
	</tr>
	<tr align="center"> 
		<td></td>
		<td bgcolor="#993300"><b><font color="#FFFFFF"><a href="http://<%=request.serverVariables("SERVER_NAME")%>/procurement/datamanager/sumprod.asp?sort=dpoDate">Date</a></font></b></td>
		<td bgcolor="#993300"><b><font color="#FFFFFF"><a href="http://<%=request.serverVariables("SERVER_NAME")%>/procurement/datamanager/sumprod.asp?sort=sProdCode">PO No.</a></font></b></td>
		<td bgcolor="#993300"><b><font color="#FFFFFF"><a href="http://<%=request.serverVariables("SERVER_NAME")%>/procurement/datamanager/sumprod.asp?sort=sProdCode">PCode</a></font></b></td>
		<td bgcolor="#993300"><b><font color="#FFFFFF"><a href="http://<%=request.serverVariables("SERVER_NAME")%>/procurement/datamanager/sumprod.asp?sort=sProdName">PName</a></font></b></td>
		<td bgcolor="#993300"><b><font color="#FFFFFF"><a href="http://<%=request.serverVariables("SERVER_NAME")%>/procurement/datamanager/sumprod.asp?sort=nQuality">Qty</a></font></b></td>
		<td bgcolor="#993300"><b><font color="#FFFFFF"><a href="http://<%=request.serverVariables("SERVER_NAME")%>/procurement/datamanager/sumprod.asp?sort=sUnitCode">Unit</a></font></b></td>
		<td></td>
	</tr>
	<%
	
	sql = "select dpodate, ppoitem.spono, sprodcode, sprodname, sum(nquality), sunitcode from ppo, ppoitem where ppo.spono = ppoitem.spono " & qryDpoDate & " " & qrySProdCode & " " & qrySProdName & " group by  dpodate, ppoitem.spono, sprodcode, sprodname, nquality, sunitcode order by " & qrySort
	'response.write sql
	'Connecting to database
	set conn = server.createObject("adodb.connection")
	set rs = server.createObject("adodb.recordset")
	conn.open "amec"
	rs.open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
	
	if not rs.EOF then
		'Set recordset page
		rs.pageSize = 50' request("selPageSize")
		pageSize = rs.pageSize
		pageCount = rs.pageCount
		pageNo = cint(request.queryString("pageNo") * 1)

		if pageNo = "" or pageNo <= 0 then
			pageNo = 1
		elseif int(pageNo) > pageCount then
			pageNo = pageCount
		end if

		rs.absolutePage = pageNo

		'Initial report variables
		dpoDate = ""
		dpoMonth = ""
		sProdCode = ""
		sProdName = ""
		sUnitCode = ""
	end if
	
	'Display data
	do while not rs.EOF%>
	<!===== Display month summary =====>
	<% 
		if dpoMonth <> month(rs("dpoDate")) and mProduct > 0 then%>
	<tr> 
		<td></td>
		<td bgcolor="#CCCCFF">&nbsp;</td>
		<td bgcolor="#CCCCFF" colspan="3"><font color="#000000"><b><%=mProduct%></b> 
			record(s) </font></td>
		<td bgcolor="#CCCCFF" align="right"><font color="#000000"><b><%=mQty%></b> 
			Unit</font></td>
		<td bgcolor="#CCCCFF">&nbsp;</td>
		<td></td>
	</tr>
	<%
			totalProduct = totalProduct + mProduct
			totalQty = totalQty + mQty
			mProduct = 0
			mQty = 0
		end if
		%>
	<tr> 
		<td>&nbsp;</td>
		<!PO date>
		<td bgcolor="#CCFFFF"><font color="#000000"> 
			<%if dpoMonth <> month(rs("dpoDate")) then
					response.write monthName(month(rs("dpoDate")), true) & "/" & year(rs("dpoDate"))
				else
					response.write "&nbsp;"
				end if%>
			</font></td>
		<td bgcolor="#CCFFFF"><font color="#000000"> 
			<%'if sProdCode <> rs("sProdCode") or dpoMonth <> month(rs("dpoDate")) then
					response.write rs("spono")
				'else
				'	response.write "&nbsp;"
				'end if%>
			</font></td>
		<!Product code>
		<td bgcolor="#CCFFFF"><font color="#000000"> 
			<%'if sProdCode <> rs("sProdCode") or dpoMonth <> month(rs("dpoDate")) then
					response.write rs("sProdCode")
				'else
				'	response.write "&nbsp;"
				'end if%>
			</font></td>
		<!Product name>
		<td bgcolor="#CCFFFF"><font color="#000000"> 
			<%'if sProdName <> rs("sProdName") or dpoMonth <> month(rs("dpoDate")) then
					response.write rs("sProdName")
				'else
				'	response.write "&nbsp;"
				'end if%>
			</font></td>
		<!nQuality>
		<td bgcolor="#CCFFFF" align="right"><font color="#000000"><%=rs("sum(nQuality)")%></font></td>
		<!Unit code>
		<td bgcolor="#CCFFFF" align="center"><font color="#000000"> 
			<%'if sUnitCode <> rs("sUnitCode") or dpoMonth <> month(rs("dpoDate")) then
					response.write rs("sUnitCode")
				'else
				'	response.write "&nbsp;"
				'end if%>
			</font></td>
		<td>&nbsp;</td>
	</tr>
	<%
		mProduct = mProduct + 1
		mQty = mQty + cint(rs("sum(nQuality)"))

		dpoDate = rs("dpoDate")
		dpoMonth = month(rs("dpoDate"))
		sProdCode = rs("sProdCode")
		sProdName = rs("sProdName")
		sUnitCode = rs("sUnitCode")
	
		rs.moveNext
		i = i + 1
		if i > 30 then
			exit do
		end if
	loop
	%>
	<tr> 
		<td></td>
		<td bgcolor="#CCCCFF">&nbsp;</td>
		<td bgcolor="#CCCCFF" colspan="3"><font color="#000000"><b><%=mProduct%></b> 
			record(s)</font></td>
		<td bgcolor="#CCCCFF" align="right"><font color="#000000"> <b><%=mQty%></b> 
			Unit</font></td>
		<td bgcolor="#CCCCFF">&nbsp;</td>
		<td></td>
	</tr>
	<%
		totalProduct = totalProduct + mProduct
		totalQty = totalQty + mQty
	%>
	<tr> 
		<td></td>
		<td bgcolor="#CC99FF">&nbsp;</td>
		<td bgcolor="#CC99FF" colspan="3"><font color="#000000">Total </font><font color="#000000"><b><%=totalProduct%></b> 
			record(s)</font></td>
		<td bgcolor="#CC99FF" align="right"><font color="#000000"> Total <b><%=totalQty%></b> 
			Unit</font></td>
		<td bgcolor="#CC99FF">&nbsp;</td>
		<td></td>
	</tr>
	<tr> 
		<td align="left" valign="bottom">&nbsp;</td>
		<td colspan="6" rowspan="2" valign="middle" align="center"><font color="#000000"> 
			<% 'Display page navigator
		if pageSize <> pageSize * pageCount then
			response.write "<a href='http://" & request.serverVariables("SERVER_NAME") & "/procurement/datamanager/sumprod.asp?pageNo=" & pageNo - 1 & "'>&lt;&lt;Previous</a>&nbsp;|&nbsp;<a href='http://" & request.serverVariables("SERVER_NAME") & "/procurement/datamanager/sumprod.asp?pageNo=" & pageNo + 1 & "'>Next&gt;&gt;</a><br>["
			
			i = 1
			'First page
			if i = pageNo then
				response.write "<b>" & pageSize & "</b>"
			else
					response.write "<a href='http://" & request.serverVariables("SERVER_NAME") & "/procurement/datamanager/sumprod.asp?pageNo=1'>" & pageSize & "</a>"
			end if
	
			for i = 2 to pageCount
				if i = pageNo then
					response.write ", <b>" & i  * pageSize & "</b>"
				else
					response.write ", <a href='http://" & request.serverVariables("SERVER_NAME") & "/procurement/datamanager/sumprod.asp?pageNo=" & i & "'>" & i  * pageSize & "</a>"
				end if
				
				if i \ 30 = 1 then
					response.write "<br>"
				end if
			next
			response.write "]"
		end if
		%>
			</font></td>
		<td align="right" valign="bottom">&nbsp;</td>
	</tr>
	<td align="left" valign="bottom"><img src="images/y03.jpg" width="12" height="12"></td>
	<td align="right" valign="bottom"><img src="images/y04.jpg" width="12" height="12"></td>
	</tr>
</table>
<%
	rs.close
	set rs = nothing
	set conn = nothing
end if
%>

<!===== Query section =====>
<form name="frmMain" method="post" action="http://<%=request.serverVariables("SERVER_NAME")%>/procurement/datamanager/sumprod.asp">
	<table cellspacing="0" cellpadding="0" border="0" bgcolor="#FFCC66" align="center">
		<tr align="center"> 
			<td align="left" valign="top"><img src="images/y01.jpg" width="12" height="12"></td>
			<td colspan="3"><b></b></td>
			<td align="right" valign="top"><img src="images/y02.jpg" width="12" height="12"></td>
		</tr>
		<tr align="center"> 
			<td align="left" valign="top">&nbsp;</td>
			<td align="left" valign="top" bgcolor="#CC6633"><b><img src="images/hCorner1.jpg" width="12" height="12"></b></td>
			<td bgcolor="#CC6633"><b><font color="#000000">PPOITEM Query</font></b></td>
			<td align="right" valign="top" bgcolor="#CC6633"><img src="images/hCorner2.jpg" width="12" height="12"></td>
			<td align="right" valign="top">&nbsp;</td>
		</tr>
		<tr> 
			<td valign="top">&nbsp;</td>
			<td valign="top" rowspan="3" colspan="3"> 
				<table border="0" cellspacing="0" cellpadding="1" width="100%">
					<tr> 
						<td valign="top" bgcolor="#993300"><b><font color="#FFFFFF">Date 
							range </font></b></td>
						<td valign="top" bgcolor="#FFCC66"> 
							<%makeSelTag m, "name=selSMonth", 0%>
							<%makeSelTag d, "name=selSDate", 0%>
							<%makeSelTag y, "name=selSYear", 0%>
							&nbsp;<font color="#000000">To</font><br>
							<%makeSelTag m, "name=selFMonth", 0%>
							<%makeSelTag d, "name=selFDate", 0%>
							<%makeSelTag y, "name=selFYear", 0%>
						</td>
					</tr>
					<tr> 
						<td bgcolor="#993300"><b><font color="#FFFFFF">Product 
							code </font></b></td>
						<td bgcolor="#FFCC66"> 
							<input type="text" name="txtProdCode" maxlength="128" size="40">
						</td>
					</tr>
					<tr> 
						<td bgcolor="#993300"><b><font color="#FFFFFF">Product 
							name</font></b></td>
						<td bgcolor="#FFCC66"> 
							<input type="text" name="txtProdName" maxlength="128" size="40">
						</td>
					</tr>
					<tr valign="middle" align="center"> 
						<td colspan="2"></td>
					</tr>
					<tr valign="middle" align="center"> 
						<td colspan="2"></td>
					</tr>
					<tr valign="middle" align="center"> 
						<td colspan="2"><b> 
							<input type="submit" value="View report">
							</b> </td>
					</tr>
					<tr> 
						<td valign="middle" colspan="2" align="center"> 
							<hr>
						</td>
					</tr>
					<tr> 
						<td bgcolor="#993300"><b><font color="#FFFFFF">Sort by</font></b></td>
						<td bgcolor="#FFCC66"> 
							<select name="selSort">
								<option value="dpoDate" selected>Date</option>
								<option value="sProdCode">Product code</option>
								<option value="sProdName">Product name</option>
								<option value="nQuality">Quantity</option>
								<option value="sUnitCode">Unit code</option>
							</select>
						</td>
					</tr>
					<tr> 
						<td bgcolor="#993300"><b><font color="#FFFFFF">Page size</font></b></td>
						<td bgcolor="#FFCC66"> 
							<select name="selPageSize">
								<option value="50" selected>50</option>
								<option value="100">100</option>
								<option value="150">150</option>
								<option value="200">200</option>
							</select>
							<font color="#000000"> records</font> </td>
					</tr>
				</table>
			</td>
			<td valign="top">&nbsp;</td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr> 
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr align="center"> 
			<td align="left" valign="bottom"><img src="images/y03.jpg" width="12" height="12"></td>
			<td valign="middle" colspan="3">&nbsp;</td>
			<td valign="bottom" align="right"><img src="images/y04.jpg" width="12" height="12"></td>
		</tr>
		<!--					<tr> 
						<td colspan="2"><b><font color="#000000">Remark: </font></b><font color="#000000">product 
							code and name are case sensitive</font> </td>
					</tr>-->
	</table>
</form>
</body>
</html>