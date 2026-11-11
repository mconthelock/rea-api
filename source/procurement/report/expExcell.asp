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
	Dim rs,sSQL
	Dim sExptype, sCattype, sCatalog, sVendcode , sCashreq , sUrgent , FromDate , ToDate
	Dim exp1, exp
	Dim rs2
	Dim i , buyer
	
	Set rs=Server.CreateObject("ADODB.Recordset")
	Set rs2=Server.CreateObject("ADODB.Recordset")
		
	    sExptype=Request("sExptype")
		sCattype=Request("sCattype")
		sCatalog=Request("sCatalog")
		sVendcode=Request("sVendcode")
		sCashreq=Request("sCashreq")
		sUrgent=Request("sUrgent")
		FromDate= Request("fromDate")
		ToDate= Request("toDate")
		
	if sExptype="0" and sCattype="0" and sCatalog="0" then 
		exp1=0
		exp=""
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog<>"0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
		exp=exp & " and "
	elseif sExptype<>"0" and sCattype<>"0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and "
	elseif sExptype<>"0" and sCattype="0" and sCatalog="0" then 
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and "
	else
		exp=" sExptype='" & sExptype & "'"
		exp=exp & " and sCattype='" & sCattype & "'"
		exp=exp & " and sCatID='" & sCatalog & "'"
		exp=exp & " and "
	end if

if sUrgent<>"1" and sCashReq<>"1" then
	if exp1="0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode ,scurrency From vPRec where to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,scurrency ,svendcode " &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode ,scurrency From vPRec where " & exp &_
		" to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency From vPRec where sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency From vPRec where " & exp &_
		" sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	end if
elseif sUrgent="1" and sCashReq<>"1" then
	if exp1="0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency From vPRec where to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename,svendcode ,scurrency From vPRec where " & exp &_
		" to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename,svendcode ,scurrency From vPRec where sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename,svendcode ,scurrency From vPRec where " & exp &_
		" sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	end if
elseif sUrgent<>"1" and sCashReq="1" then
	if exp1="0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency From vPRec where to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency From vPRec where " & exp &_
		" to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename,svendcode ,scurrency From vPRec where sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency From vPRec where " & exp &_
		" sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	end if
elseif sUrgent="1" and sCashReq="1" then
	if exp1="0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency From vPRec where to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode="0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename,svendcode ,scurrency From vPRec where " & exp &_
		" to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp1="0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename,svendcode ,scurrency From vPRec where sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,scurrency,svendcode" &_
		" order By sRecno,sRecby,sCurcode"
	elseif exp<>"0" and sVendcode<>"0" then
		sSQL="Select sCurcode,Sum(nAmount) as SumAmount,sRecno,sRecby,sprno,spono,sename ,svendcode,scurrency  From vPRec where " & exp &_
		" sVendcode='" & svendcode & "'" &_
		" and to_date(dRecdate) >=" & To_Date(FromDate) &_
		" and to_date(dRecdate) <=" & To_Date(ToDate) &_
		" and cPrtype='1'" &_
		" and sCashReq<>' '" &_
		" Group By sCurcode,sRecno,sRecby,sprno,spono,sename,svendcode,scurrency" &_
		" order By sRecno,sRecby,sCurcode"
	end if
End if
		
	'response.write sSQL
	'response.end
	sSQL = "select * From ("&sSQL&") V LEFT OUTER JOIN  (select sempno as buyer , sprno from amec.pprapp where amec.pprapp.nstep = '35') M  ON V.sprno = M.sprno ORDER BY M.buyer , V.spono , V.sename , V.svendcode "
	'response.write sSQL
	'response.end
	
	rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	
%>
<HTML xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<title>Search Result</title>
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
<center><font color="#000000" size="6">Search By <%=FromDate&" - "&ToDate%></font></center>
<table width="100%" border="1" cellspacing="1" cellpadding="3" bgColor="#dedede">
  <tr bgColor="#f6f6f6"> 
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>No.</b></font></td>  
	 <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Vendor code</b></font></td>  
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Vendor name</b></font></td>    	
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Receive No.</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>PO No.</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Buyer</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Amount</b></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="6"><b>Currency</b></font></td>
  </tr>
<%	
		i = 0
		do while not rs.eof
			i = i+1
			'sSQL = "select p.sempno from  amec.pprapp p where  p.sprno = '"&rs("sprno")&"' and p.nstep = '35' "
			
			'rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
			'buyer = ""
			'if Not rs2.EoF then
			'	buyer = rs2("sempno")
			'end if
			'rs2.close
%>
  <tr bgcolor="#ffffff"> 
  <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=i%></font></td>
	<td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("svendcode")%></font></td>	
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sename")%></font></td>	
	<td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sRecby")%></font></td>
	<td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("spono")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("buyer")%></font></td>	
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("SumAmount")%></font></td>
    <td nowrap align="center"><font face="AngsanaUPC" size="5"><%=rs("sCurcode")%></font></td>
  </tr>
<%			
					rs.MoveNext
			Loop
			rs.close
%>
</table>
</body>
</html>
