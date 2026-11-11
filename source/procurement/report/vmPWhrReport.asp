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
<%
	Dim RS,sSQL
	Dim sExptype,From_Day,From_Month,From_Year,To_Day,To_Month,To_Year,FromDate,ToDate
	Dim svendcode,sCattype,sCatalog,exp,exp1
	sExptype=Request("ListProduct")
	sCattype=Request("ListCattype")
	sCatalog=Request("ListCatalog")
	sVendcode=Request("vender")
	if Trim(Request("Vender"))="" then sVendcode="0"
	From_Day="1"
	From_month=Request("mFrom_Month")
	From_Year=Request("mFrom_Year")
''		FromDate=From_Month & "/" & From_Day & "/" & From_Year
	FromDate=From_Day & "/" & From_Month & "/" & From_Year
	To_month=Request("mTo_Month")
	To_Year=Request("mTo_Year")
	To_Day=MonthDay(To_Month,To_Year)
''		ToDate=To_Month & "/" & To_Day & "/" & To_Year
	ToDate=To_Day & "/" & To_Month & "/" & To_Year
	if sExptype="0" and sCattype="0" and sCatalog="0" then 
		exp1=0
		exp=""
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog<>"0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
	elseif sExptype<>"0" and sCattype="0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
	else
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
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
<HTML xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<title>WHR Summary report</title>
<meta http-equiv="Content-Type" content="text/html; charset=tis-620">
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
			<td  align="center" bgcolor="#f0f0f5">PR No.</td>
			<td  align="center" bgcolor="#f0f0f5">PO No.</td>
			<td  align="center" bgcolor="#f0f0f5">Name</td>
			<td  align="center" bgcolor="#f0f0f5">Section</td>
			<td  align="center" bgcolor="#f0f0f5">Department</td>
			<td  align="center" bgcolor="#f0f0f5">Product Code</td>
			<td  align="center" bgcolor="#f0f0f5">Product Name</td>
			<td  align="center" bgcolor="#f0f0f5">Quantity</td>
			<td  align="center" bgcolor="#f0f0f5">Unit</td>
			<td  align="center" bgcolor="#f0f0f5">Month Receive</td>
			<td  align="center" bgcolor="#f0f0f5">Invoice</td>
			<td  align="center" bgcolor="#f0f0f5">Expense Type</td>
			<td  align="center" bgcolor="#f0f0f5">Group in Exp</td>
			<td  align="center" bgcolor="#f0f0f5">Sub-Group in Exp</td>
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
			<td  valign="top"><%=RS("srecby")%></td>
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