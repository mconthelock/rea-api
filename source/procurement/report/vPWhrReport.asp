<% Option Explicit %>
<%
Response.ContentType ="Application/vnd.excel" 
Response.AddHeader "Content-Disposition", "attachment; filename=whrsummaryreport"& day(date) &"-"& month(date) & "-" & year(date) & ".xls" 
%>
<%
           ' Author : Ekkapan Kongchoo
%>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<HTML xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<title>WHR Summary report</title>
<meta http-equiv="Content-Type" content="text/html; charset=tis-620">
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
<%
	Dim RS,sSQL
	Dim sExptype,From_Day,From_Month,From_Year,To_Day,To_Month,To_Year,FromDate,ToDate
	Dim svendcode,sCattype,sCatalog,exp,exp1
	sExptype=Request("ListProduct")
	sCattype=Request("ListCattype")
	sCatalog=Request("ListCatalog")
	sVendcode=Request("vender")
	if Trim(Request("Vender"))="" then sVendcode="0"
	
	From_Day=Request("From_Day")
	From_month=Request("From_Month")
	From_Year=Request("From_Year")
	FromDate= From_Day & "/" & From_Month & "/" & From_Year
	To_Day=Request("To_Day")
	To_month=Request("To_Month")
	To_Year=Request("To_Year")
	ToDate=To_Day & "/" & To_Month & "/" & To_Year
	if sExptype="0" and sCattype="0" and sCatalog="0" then 
		exp1=0
		exp=""
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog<>"0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
		'exp=exp & " and "
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		'exp=exp & " and "
	elseif sExptype<>"0" and sCattype="0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
		'exp=exp & " and "
	else
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
		'exp=exp & " and "
	end if
	
	sSQL = "select sprno , spono , sname , ssec , sdept , sprodcode , sprodname , nquality , sunitcode , drecdate as monrec  , sexpdesc , scatdesc , scatalog , srecby from vPRec"
	sSQL = sSQL&" where to_date(dRecdate) >=" & To_Date(FromDate) &" and  to_date(dRecdate) <= " & To_Date(ToDate)
	if sVendcode <> "0" then
			sSQL = sSQL & " and svendcode = '"&sVendcode&"'"
	end if
	if exp <> "" then
			sSQL = sSQL &" and "&exp
	end if
    sSQL = sSQL & " order by drecdate"
	Set RS=Server.CreateObject("ADODB.Recordset")
	RS.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText

%>
</head>
<body>
	<table border=1 align="Center" cellspacing="1" cellpadding="0" height="18">
	  <tr>
			<td  align="center" bgcolor="#f0f0f5"><b>PR No.</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>PO No.</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Name</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Section</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Department</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Product Code</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Product Name</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Quantity</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Unit</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Month Receive</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Invoice</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Expense Type</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Group in Exp</b></td>
			<td  align="center" bgcolor="#f0f0f5"><b>Sub-Group in Exp</b></td>
	  </tr>
     <%
		if RS.EOF then 
	  %>
	    <tr><td colspan="13"  align="center"><font color="red">Data not found</font></td></tr>
	  
	  
	  <%
	  else
		Do While Not RS.EOF 
	  %>
		<tr>
			<td  valign="top"><%=RS("sprno")%></td>
			<td  valign="top"><%=RS("spono")%></td>
			<td  valign="top"><%=RS("sname")%></td>
			<td  valign="top"><%=RS("ssec")%></td>
			<td  valign="top"><%=RS("sdept")%></td>
			<td  valign="top"><%=RS("sprodcode")%></td>
			<td  valign="top"><%=RS("sprodname")%></td>
			<td  align="center" valign="top"><%=RS("nquality")%></td>
			<td  valign="top"><%=RS("sunitcode")%></td>
			<td  align="center" valign="top"><%=MONTH(RS("monrec"))&"-"&YEAR(RS("monrec")) %></td>
			<td valign="top"><%=RS("srecby")%></td>
			<td  valign="top"><%=RS("sexpdesc")%></td>
			<td  valign="top"><%=RS("scatdesc")%></td>
			<td  valign="top"><%=RS("scatalog")%></td>
		</tr>
	  <%
			RS.MoveNext
			Loop 
		end if
	  %>
	 </table>
</body>
</html>