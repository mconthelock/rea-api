<% Option Explicit %>
<%
Response.ContentType ="Application/vnd.excel" 
Response.AddHeader "Content-Disposition", "attachment; filename=product-"& request("sExptype") & "-" &  day(date) &"-"& month(date) & "-" & year(date) & ".xls" 
%>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<% 
	Dim rs,sSQL
	Dim id
	Set rs=Server.CreateObject("ADODB.Recordset")
	id = 1
	sSQL= request("sSQL")
	
	
	
	rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
%>
<HTML xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<title>Search Resault</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
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
<!--<center><font color="#000000" size="6">Search By <%'= request("sExptype")%></font></center>-->
<table width="100%" border="1" cellspacing="1" cellpadding="3" bgColor="#dedede">
  <tr><td></td></tr>
  <tr bgColor="#f6f6f6"> 
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>No.</b></font></td>    
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Product Code</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Product Name</b></font></td>	
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Expense Type</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Group in Expense Type</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Sub-Group in Expense Type</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>UPrice</b></font></td>
  </tr>
<%	
		do while not rs.eof  
%>
  <tr bgcolor="#ffffff"> 
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=id%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sprodcode")%></font></td>	
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("seprodname")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sExptype")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sCattype")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sCatID")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=FormatNumber(RS("nPrice"), 3)%></font></td>
  </tr>
<%			
					rs.MoveNext
					id = id + 1
				Loop
				rs.close
%>
</table>
</body>
</html>