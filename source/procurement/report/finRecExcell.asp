<% Option Explicit %>
<%
Response.ContentType ="Application/vnd.excel" 
Response.AddHeader "Content-Disposition", "attachment; filename=receive"& request("sExptype") & day(date) &"-"& month(date) & "-" & year(date) & ".xls" 
%>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	Dim RS,sSQL
	Dim sSQL2,RS2, whereClause
	Dim cntLine, Color, RowCount
	cntLine = 1
	Set RS2=Server.CreateObject("ADODB.Recordset")

	Color="#FFCCCC"

	sSQL = request("sSQL")
	Set RS=Server.CreateObject("ADODB.Recordset")
	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<title>Search Resault</title>
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
	<table border=1 align="Center" cellspacing="1" cellpadding="0" height="18">
	  <tr>
		<td height="20" align="center"> No.</td>        
		<td height="20" align="center">PR No.</td>
		<td height="20" align="center">PR Date</td>
		<td height="20" align="center">Requester</td>
		<td height="20" align="center">Amount (PR)</td>
		<td height="20" align="center">Currency</td>
		<td height="20" align="center">Vendor Name</td>
		<td height="20" align="center">Expense Type</td>
		<td height="20" align="center">Group in Exp</td>
		<td height="20" align="center">Sub Group in Exp</td>
		<td height="20"  align="center">Catalog</td>
		<td height="20" align="center">PO No.</td>
		<td height="20" align="center">PO Date</td>
		<td height="20" align="center">Amount (PO)</td>
		<td height="20" align="center">Invoice No.</td>						
		<td height="20" align="center">Receive Date</td>
		<td height="20" align="center">Amount (Invoice)</td>						
	  </tr>
		<%
		If Not Rs.EOF then
			'RS.Absolutepage=PageNo
			RowCount=1
			'Do While Not Rs.EOF and RowCount<=PageSize
			Do While Not Rs.EOF 
			%>
		  <tr> 
			<td height="20" bgcolor="<%=Color%>"><%=RowCount%>&nbsp;</td>
			<td height="20" bgcolor="<%=Color%>"><%=Rs("sPrNo")%></td>
			<td height="20" bgcolor="<%=Color%>"><%=RS("dPrDate")%></td>
			<td height="20" bgcolor="<%=Color%>"><%=RS("sRequester")%></td>
			<td height="20" bgcolor="<%=Color%>" align="right"><%=FormatNumber(RS("prAmn"), 4)%></td>
			<td height="20" bgcolor="<%=Color%>">&nbsp;<%=RS("sCurcode")%></td>
			<td height="20" bgcolor="<%=Color%>">&nbsp;<%=RS("sEname")%></td>
			<td height="20" bgcolor="<%=Color%>">&nbsp;<%=RS("sExptype")%></td>												
			<td height="20" bgcolor="<%=Color%>">&nbsp;<%=RS("sCatid")%></td>																		
			<td height="20" bgcolor="<%=Color%>">&nbsp;<%=RS("sCattype")%></td>
			<td height="20" bgcolor="<%=Color%>" >&nbsp;<%=RS("scatalog")%></td>						
			<td height="20" bgcolor="<%=Color%>">&nbsp;<%=RS("sPoNo")%></td>												
			<td height="20" bgcolor="<%=Color%>"><%=RS("dPodate")%></td>
			<td height="20" bgcolor="<%=Color%>" align="right"><%=FormatNumber(RS("poAmn"), 4)%></td>
			<%
			sSQL = "select srecby, to_char(drecdate,'dd/MM/yyyy') drecdate , recAmount from vuserRec where spono = '" & RS("sPoNo") & "' and scatid = '" & RS("sCatid") & "' and scattype = '"& RS("sCattype") & "' and sexptype = '"& RS("sExptype") &"'" 
		'	response.write sSQL
			Rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
			if Not Rs2.EOF then
				Do While Not RS2.EOF
			%>
				<td height="20" bgcolor="<%=Color%>">&nbsp;<%=RS2("srecby")%></td>
				<td height="20" bgcolor="<%=Color%>"><%=RS2("drecdate")%></td>
				<td height="20" bgcolor="<%=Color%>" align="right" width="50%"><font color="blue"><%=FormatNumber(RS2("recAmount"), 4)%></font></td>		
			<%			
					RS2.MoveNext
					cntLine = cntLine +1 
					if (cntLine > 1) and (Not RS2.EoF) then
			%>
		</tr>
		<tr>
		  <td height="20" colspan=14 bgcolor="<%=Color%>">&nbsp;</td>
		<%	
					end if
				Loop
		Else
			response.write "<td height=20 bgcolor=" & Color &">&nbsp;</td>"
			response.write "<td height=20 bgcolor=" & Color &">&nbsp;</td>"
			response.write "<td height=20 bgcolor=" & Color &">&nbsp;</td>"												
		End if
		Rs2.Close
		%>
		</tr>
		<%
		if Color="#FFCCCC" then
			Color="#99CCFF"
		else
			Color="#FFCCCC"
		end if
		RowCount=RowCount+1
		RS.MoveNext
		Loop 
		End if
%>
  </table>  
</BODY>
</HTML>
