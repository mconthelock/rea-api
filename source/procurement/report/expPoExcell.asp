<% Option Explicit %>
<%
Response.ContentType ="Application/vnd.excel" 
Response.AddHeader "Content-Disposition", "attachment; filename=po-"& request("sExptype") & "-" &  day(date) &"-"& month(date) & "-" & year(date) & ".xls" 
%>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<% 
	Dim rs,sSQL
	Set rs=Server.CreateObject("ADODB.Recordset")

		sSQL="select spono, srequester, sprodname, nquality," _
		& " nprice,namount, scurrency, sprodcode " _
		& " from vrPPO" _
		& " where " _
		& " dPOdate >= " & To_Date(request("fromDate"))  _
		& " and dPOdate <= " & To_Date(request("toDate")) _
		& " and cstatus<>'0' " _
		& " and cstatus<>'9' "
		if request("sExptype") = "51" then
'			sSQL = sSQL & " and sRequester='92222'"
		end if
		if request("sExptype") = "52" then
'			sSQL = sSQL & " and sRequester='92222'"		
		end if
'		if request("sExptype") <> "0" then
'			sSQL = sSQL & " and sExptype ='" & request("sExptype") & "'"		
'		end if
		if request("sExptype") <> "0" then
		sSQL = sSQL & " and sExptype ='" & request("sExptype") & "'"
		end if
		if request("sCattype") <> "0" then
			sSQL = sSQL & " and sCattype ='" & request("sCattype") & "'"		
		end if
		if request("sCatalog") <> "0" then
			sSQL = sSQL & " and sCatId ='" & request("sCatalog") & "'"
		end if
		if request("sVendcode") <> "0" then
			sSQL = sSQL & " and sVendcode = '" & request("sVendcode") & "'"
		end if		
'		& " and sExptype='" & request("sExptype") & "' and dRecdate >= to_date('" & request("fromDate") & "','MM/DD/YYYY')" _
'		& " and dRecdate <= to_date('" & request("toDate") & "','MM/DD/YYYY') " 
'		response.write sSQL
'		response.end
		rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
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
<center><font color="#000000" size="6">Search By <%= request("sExptype")%></font></center>
<table width="100%" border="1" cellspacing="1" cellpadding="3" bgColor="#dedede">
  <tr bgColor="#f6f6f6"> 
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>PO No.</b></font></td>    
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Requester</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Product Code</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Product Name</b></font></td>	
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Quantity</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Price</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Amount</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Currency</b></font></td>
  </tr>
<%	
		do while not rs.eof  
%>
  <tr bgcolor="#ffffff"> 
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("spono")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("srequester")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sprodcode")%></font></td>	
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sprodname")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("nquality")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("nprice")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("namount")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("scurrency")%></font></td>
  </tr>
<%			
					rs.MoveNext
				Loop
				rs.close
%>
</table>
</body>
</html>