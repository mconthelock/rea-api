<% Option Explicit %>
<%
Response.ContentType ="Application/vnd.excel" 
Response.AddHeader "Content-Disposition", "attachment; filename=planexp" & day(date) &"-"& month(date) & "-" & year(date) & ".xls" 
%>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	Dim rs,sSQL,whereClause , i
	Set rs=Server.CreateObject("ADODB.Recordset")

	'########## GENERATE WHERE CLAUSE ##########
										whereClause = "ppoitem.spono = ppo.spono and ppoitem.sprodcode = planreceiveitem.sprodcode " _
										& "and planreceive.spono = ppo.spono and planreceive.ID = planreceiveitem.ID and ppo.srequester = aemployee.sempno " _
										& "and aemployee.sseccode = psection.sseccode and aemployee.sdepcode = pdepartment.sdepcode and aemployee.sdivcode = pdivision.sdivcode " _
										& "and ppo.cstatus <> '9' and (ppo.cstatus = '2' or ppo.cstatus = '3') "
										
										'Po number
										if request.queryString("div") <> "0" then
											whereClause = whereClause & " and aemployee.sdivcode = '" & request.queryString("div") & "'"
										end if
										
										'Expense type
										if request.queryString("dept") <> "0" then
											whereClause = whereClause & " and aemployee.sdepcode  = '" & request.queryString("dept") & "'"
										end if
										
										'Vendor
										if request.queryString("sec") <> "0" then
											whereClause = whereClause & " and aemployee.sseccode = '" & request.queryString("sec") & "'"
										end if
										
										'Issue date (from)
										if request.queryString("fromDate") <> "" then
											whereClause = whereClause & " and precDate >= to_date('" & request.queryString("fromDate") & "', 'DD/MM/YYYY')"
										end if

										'Issue date (to)
										if request.queryString("toDate") <> "" then
											whereClause = whereClause & " and precDate <= to_date('" & request.queryString("toDate") & "', 'DD/MM/YYYY')"
										end if
								'########## END GENERATE WHERE CLAUSE ##########
										
										'########## GOTO START RECORD ##########
										sSQL = "select precdate , prectime , precplace , planreceive.spono , planreceiveitem.sprodcode , planreceiveitem.sprodname , planreceiveitem.nquality , planreceiveitem.sunitcode" _
										& ", aemployee.sname , psection.ssec , pdepartment.sdept , pdivision.sdiv" _
										& " from ppo, ppoItem, aemployee , planreceive , planreceiveitem , psection , pdepartment , pdivision " _
										& "where " & whereClause _
										& " order by precdate, prectime , planreceive.sPoNo, planreceiveitem.sProdName "
										rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
										
%>
<HTML xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<title>Plan Receive Result</title>
<meta http-equiv="Content-Type" content="text/html; charset=">
<META http-equiv=Content-Type content="text/html; charset=us-ascii"> 
<META name=ProgId content=Excel.Sheet> 
<META name=Generator content="Microsoft Excel 9"> 
<LINK rel=File-List href="./z.files/filelist.xml"> 
<LINK rel=Edit-Time-Data href="./z.files/editdata.mso"> 
<LINK rel=OLE-Object-Data href="./z.files/oledata.mso"> 

<!--[if gte mso 9]>
<xml> 
<o:DocumentProperties> 
<o:LastAuthor>wonder</o:LastAuthor> 
<o:LastPrinted>2003-05-08T07:02:15Z</o:LastPrinted> 
<o:Created>2003-05-07T10:59:26Z</o:Created> 
<o:LastSaved>2003-05-08T07:02:18Z</o:LastSaved> 
<o:Version>9.6615</o:Version> 
</o:DocumentProperties> 
<o:OfficeDocumentSettings> 
<o:DownloadComponents/> 
<o:LocationOfComponents HRef="file:///D:/program/application/office%25202000/msowc.cab"/> 
</o:OfficeDocumentSettings> 
</xml>
<![endif]--> 

</head>
<body>
<center><font color="#000000" size="6">Plan Receive <%=request.queryString("fromDate")&"-"&request.queryString("toDate")%></font></center>
<table width="100%" border="1" cellspacing="1" cellpadding="3" bgColor="#dedede">
  <tr bgColor="#f6f6f6"> 
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
			i = 0
		do while not rs.eof  
	
%>
  <tr bgcolor="#ffffff"> 
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=i+1%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"> <%=rs("precdate")%></font></td>	
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("prectime")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("precplace")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("spono")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sprodcode")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sprodname")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("nquality")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sunitcode")%></font></td>
	<td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sname")%></font></td>
	<td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("ssec")%></font></td>
	<td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sdept")%></font></td>
	<td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sdiv")%></font></td>
  </tr>
<%			
	i = i+1
	rs.MoveNext		
	Loop
	rs.close
%>
</table>
</body>
</html>