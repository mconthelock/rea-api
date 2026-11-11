<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../List.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: worakrit yodthonglert
	' Version: 1.01
%>
<%
		Dim Rs,sSQL,Rs2,sSQL2,Rs3,sSQL3,Rs4,sSQL4,Rs5,sSQL5,Rs6,sSQL6,Rs7,sSQL7,Rs8,sSQL8
		Dim SumArr(),CattypeArr()
		Dim Color
		Dim nItem,nItemCount,nAmount,LastnAmount,TotalAmount
		Dim CatIDCount,ProductCount,CattypeCount,sExptype
		Dim Counter,nCount,sCount,TotalPrice,sCount2,LastItem
		Dim FormDate,ToDate
		Dim PrCount,PoCount
		Dim SummaryBY,Message
%>
<%
		Color = "#FFFFFF"
		nItem = 0
		TotalAmount = 0
		nItemCount = 0
		nAmount = 0
		LastnAmount = 0
		Counter = 0
		nCount = 0
		TotalPrice = 0
		LastItem = 0
		sCount = 0
%>
<%
		sExptype = Request("sExptype")
		SummaryBY = Request("SummaryBY")
		FormDate = Request("Form_Month")&"/"&Request("Form_Day")&"/"&Request("Form_Year")
		ToDate = Request("To_Month")&"/"&Request("To_Day")&"/"&Request("To_Year")
%>
<%

		sSQL = "Select * From PCATTYPE where sExptype = '"&sExptype&"' Order by sCattype"
		Set Rs = CreateObject("ADODB.Recordset")
		Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
		CattypeCount = Rs.RecordCount
		Redim CattypeArr(CattypeCount)
		While Not Rs.EOF
		sCount = sCount + 1
			CattypeArr(sCount) = Rs("sCattype")
		Rs.MoveNext
		Wend
	sCount = 0
	if CattypeCount = 0 Then
	else
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@
					'@@@@@@@@@         Comment      @@@@@@@@@@@
					'	(*,0) = P/R No.
					'	SumArr(*,1) = P/O No.
					'	SumArr(*,2) = Expensetype Code
					'	SumArr(*,3) = Currency
					'	SumArr(*,4) = Receive Total Price
					'	SumArr(*,5) = Invoice Total Price
					'	SumArr(*,6) = Tax Invoice Total Price
					'	SumArr(*,7) = P/O Item Total Price
					'	SumArr(*,8) = P/R Item Total Price
					'	SumArr(*,9) = Catalog Type
					'	SumArr(*,10) = Catalog ID
					'	SumArr(*,11) = Product Code
					'	SumArr(*,12) = Product Quality
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@

		sSQL8 = "Select * From ppr"
		Set Rs8 = CreateObject("ADODB.Recordset")
		Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText		
		PrCount = Rs8.RecordCount 

		sSQL6 = "Select * From PCATALOG Order by sCattype"
		Set Rs6 = CreateObject("ADODB.Recordset")
		Rs6.Open sSQL6, Conn, adOpenStatic, adLockReadOnly, adCmdText
		CatIDCount = Rs6.RecordCount
		While Not Rs6.EOF
		sCount = sCount + 1 

			sSQL3 = "Select * From pProduct"
			Set Rs3 = CreateObject("ADODB.Recordset")
			Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
			ProductCount = Rs3.RecordCount
			ReDim SumArr(ProductCount,20)
				sSQL3 = "Select * From pProduct where sCatID='"&Rs6("sCatID")&"'"
				Set Rs3 = CreateObject("ADODB.Recordset")
				Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
				ProductCount = Rs3.RecordCount
					if ProductCount = 0 Then
							SumArr(nCount,11) = ""
							SumArr(nCount,10) = ""
							SumArr(nCount,9) = Rs6("sCattype")
							SumArr(nCount,2) = sExptype
'							Response.write " nCount_loop1 = "&nCount&"<BR>"
					else
							Do While Not ProductCount = 0
								While Not  Rs3.EOF
									if ProductCount = 0 Then 
									else 	nCount = nCount + 1
									end if
									SumArr(nCount,11) = Rs3("sProdcode")
									SumArr(nCount,10) = Rs6("sCatID")
									SumArr(nCount,9) = Rs6("sCattype")
									SumArr(nCount,2) = sExptype
									ProductCount = ProductCount - 1
'									Response.write "<BR> nCount_loop2 = "&nCount
								Rs3.MoveNext
								Wend	
							Loop
					end if

		Rs6.MoveNext
		Wend

		Counter = nCount
		Get_PRItem()
		Get_PoItem()	
		Get_Receive()
		Get_Invoice()
		Get_Tax()	
end if
%>

<HTML>
<HEAD>
<TITLE>Summary By Date</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>

<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="2%">&nbsp;</td>
    <td width="96%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
      </font></b></td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="2%">&nbsp;</td>
    <td width="96%" valign="TOP"> 
      <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <tr> 
          <td height="11" colspan="11"> 
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#3399CC"> 
                <td> <b><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;Summary 
                  Report By Group : <%=Message%></font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td colspan="4" bgcolor="#33CCFF"> 
            <div align="left">&nbsp;&nbsp;&nbsp;&nbsp;<b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Expense 
              Type : <%pExptype_Desc(sExptype)%></font></b></div>
          </td>
          <td width="4%">&nbsp;</td>
          <td width="8%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
        </tr>
        <tr> 
          <td bgcolor="#33CCFF" width="6%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Group 
              Code</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="10%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">GroupName</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="11%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Product 
              Code</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="11%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Product 
              Name</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="4%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">QTY</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="8%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Currency</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="10%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">P/R 
              AMT</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="10%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">P/O 
              AMT</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="10%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Receive 
              AMT</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="10%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">INV 
              AMT</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="10%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">TAX 
              INV AMT</font></b></div>
          </td>
        </tr>
        <% 
For nCount = 1 to CattypeCount
	For sCount = 1 to Counter
		if CattypeArr(nCount) = SumArr(sCount,9) Then
			if SumArr(sCount,8) <> "" Then
%> 
        <tr> 
          <td bgcolor='<%=Color%>' width="6%" height="10" align="center"> <%=SumArr(sCount,9)%> 
            </td>
          <td bgcolor='<%=Color%>' width="10%" height="10" align="center"><font color="red"><%'=sCount%></font> <%Show_CatalogDES(SumArr(sCount,9))%> 
            </td>
          <td bgcolor='<%=Color%>' width="11%" height="10" align="center"> <%=SumArr(sCount,11)%> 
            </td>
          <td bgcolor='<%=Color%>' width="11%" height="10" align="center"> <%Show_Product(SumArr(sCount,11))%> 
            </td>
          <td bgcolor='<%=Color%>' width="4%" height="10" align="center"> <%=SumArr(sCount,12)%> 
            </td>
          <td bgcolor='<%=Color%>' width="8%" height="10" align="center"> <%'=SumArr(sCount,3)%> <%Show_Currency(SumArr(sCount,3))%> 
            </td>
          <td bgcolor='<%=Color%>' width="10%" height="10" align="right"> <%=FormatNumber(SumArr(sCount,8), 3)%> 
            &nbsp;</td>
          <td bgcolor='<%=Color%>' width="10%" height="10" align="right"> <%=FormatNumber(SumArr(sCount,7), 3)%> 
            &nbsp;</td>
          <td bgcolor='<%=Color%>' width="10%" height="10" align="right"> <%=FormatNumber(SumArr(sCount,4), 3)%> 
            &nbsp;</td>
          <td bgcolor='<%=Color%>' width="10%" height="10" align="right"> <%=FormatNumber(SumArr(sCount,5), 3)%> 
            &nbsp;</td>
          <td bgcolor='<%=Color%>' width="10%" height="10" align="right"> <%=FormatNumber(SumArr(sCount,6), 3)%> 
            &nbsp;</td>
        </tr>
        <%
if Color="#99CCFF" then
		Color="#FFFFFF"
else
		Color="#99CCFF"
end if
%>
<%
			end if
		if sCount = Counter Then 
%>
        <tr> 
          <td width="6%">&nbsp;</td>
          <td valign="BOTTOM" width="10%">&nbsp; </td>
          <td valign="BOTTOM" colspan="9"> 
            <hr size="1">
          </td>
        </tr>
        <tr> 
          <td width="6%"> 
            <div align="CENTER"></div>
          </td>
          <td width="10%">&nbsp;</td>
          <td colspan="3"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">SUB 
              TOTAL</font></b></div>
          </td>
          <td width="8%" bgcolor="#99CCFF" Align="Center">  <%Show_Currency(SumArr(sCount,3))%></td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(8)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(7)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(4)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(5)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(6)%>&nbsp;</td>
        </tr>
 <% 
		end if
		end if
	Next
Next 
%> 
        <tr> 
          <td width="6%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="11%">&nbsp;</td>
          <td width="11%">&nbsp;</td>
          <td width="4%">&nbsp;</td>
          <td width="8%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="6%" height="16">&nbsp;</td>
          <td width="10%" height="16">&nbsp;</td>
          <td colspan="3" height="16"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">GRAND 
              TOTAL</font></b></div>
          </td>
          <td width="8%" bgcolor="#99CCFF" height="16" ALIGN="CENTER"><%if counter <> 0 Then Show_Currency(SumArr(sCount,3))%></td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if counter <> 0 Then SumGrandTotalPrice(8)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if counter <> 0 Then SumGrandTotalPrice(7)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if counter <> 0 Then SumGrandTotalPrice(4)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if counter <> 0 Then SumGrandTotalPrice(5)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if counter <> 0 Then SumGrandTotalPrice(6)%>&nbsp;</td>
        </tr>
        <tr> 
          <td width="6%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="11%" height="14">&nbsp;</td>
          <td width="11%" height="14">&nbsp;</td>
          <td width="4%" height="14">&nbsp;</td>
          <td width="8%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
        </tr>
      </table>
      <br>
      <hr size="1">
    </td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="2%">&nbsp;</td>
    <td width="96%">&nbsp;</td>
    <td width="2%">&nbsp;</td>
  </tr>
</table>
</BODY>
</HTML>

<%
Function Get_PoItem()

		For sCount = 1 to Counter 
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		nCount = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL7 = "Select Distinct(pPoitem.sProdcode),pPoitem.nAmount,pPr.cPo,pPo.dPoDate From pPr,pPo,pPoitem where "&_
						" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and pPo.sPono=ppoitem.sPono) "&_
						" and (pPr.cPo = 1 and pPo.dPodate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&") Order by pPo.dPodate" 
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL7 = "Select Distinct(pPoitem.sProdcode),pPoitem.nAmount,pPr.cPo,pPo.dPoDate From pPr,pPo,pPoitem where "&_
						" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and pPo.sPono=ppoitem.sPono) "&_
						" and (pPr.cPo = 1 and pPo.dPodate >= "&To_Date(FormDate)&") Order by pPo.dPodate" 
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL7 = "Select Distinct(pPoitem.sProdcode),pPoitem.nAmount,pPr.cPo,pPo.dPoDate From pPr,pPo,pPoitem where "&_
						" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and pPo.sPono=ppoitem.sPono) "&_
						" and (pPr.cPo = 1 and pPo.dPodate <= "&To_Date(ToDate)&") Order by pPo.dPodate" 
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL7 = "Select Distinct(pPoitem.sProdcode),pPoitem.nAmount,pPr.cPo,pPo.dPoDate From pPr,pPo,pPoitem where "&_
						" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and pPo.sPono=ppoitem.sPono) "&_
						" and pPr.cPo = 1 " 
		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'			Response.write sSQL7&"<BR>"
			Set Rs7 = CreateObject("ADODB.Recordset")	
			Rs7.Open sSQL7, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs7.RecordCount

'			Response.write " nItemCount = "&nItemCount&"<BR>"			
			While Not Rs7.EOF
				nAmount = cDbl(Rs7("nAmount"))
				SumArr(sCount,7) = cDbl(nAmount) + cDbl(LastnAmount)
'				Response.write " Poitem = "&SumArr(sCount,7)
'				Response.write " nAmount = "&nAmount
'				Response.write " nItemCount = "&nItemCount
'				Response.write " LastnAmount = "&LastnAmount
'				Response.write "<BR>"
				LastnAmount = SumArr(sCount,7)
			Rs7.MoveNext
			Wend

		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Next

End Function
%>

<%
Function Get_PRItem()

		nCount = 1
		For sCount = 1 to Counter 
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
			sSQL8 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,ppr Where "&_
							" pPritem.sProdcode = '"&SumArr(sCount,11)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" pPr.dPrdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&" and "&_
							" (pPritem.cStatus = 1 or pPritem.cStatus = 2) "&_
							" Order by pPr.dPrdate"
			Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
			sSQL8 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,ppr Where "&_
							" pPritem.sProdcode = '"&SumArr(sCount,11)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" pPr.dPrdate >= "&To_Date(FormDate)&" and "&_
							" (pPritem.cStatus = 1 or pPritem.cStatus = 2) "&_
							" Order by pPr.dPrdate"
			Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
			sSQL8 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,ppr Where "&_
							" pPritem.sProdcode = '"&SumArr(sCount,11)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" pPr.dPrdate <= "&To_Date(ToDate)&" and "&_
							" (pPritem.cStatus = 1 or pPritem.cStatus = 2) "&_
							" Order by pPr.dPrdate"
			Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
			sSQL8 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,ppr Where "&_
							" pPritem.sProdcode = '"&SumArr(sCount,11)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" (pPritem.cStatus = 1 or pPritem.cStatus = 2)"
			Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'			Response.write sSQL8&"<BR>"
			Set Rs8 = CreateObject("ADODB.Recordset")	
			Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs8.RecordCount
			While Not Rs8.EOF
				SumArr(sCount,12) = Rs8("nQuality")
				SumArr(sCount,8) = Rs8("nAmount")
				SumArr(sCount,3) = Rs8("sCurcode")
'				Response.write " Prodcode = "&SumArr(sCount,11)
'				Response.write " sPrno = "&Rs8("sPrno")&"<BR>"

				nAmount = cDbl(Rs8("nAmount"))
				SumArr(sCount,8) = cDbl(nAmount) + cDbl(LastnAmount)
'				Response.write " Prno = "&SumArr(sCount,0)
'				Response.write " Pritem = "&SumArr(sCount,8)
'				Response.write " nAmount = "&nAmount
'				Response.write " LastnAmount = "&LastnAmount
'				Response.write "<BR>"
				LastnAmount = SumArr(sCount,8)
				SumArr(sCount,3) = Rs8("sCurcode")
			nCount = nCount + 1
			if nCount = (nItemCount+0) Then nCount = 0
			Rs8.MoveNext
			Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Next

End Function
%>

<%
Function Get_InVoice()
	For sCount = 1 to Counter 

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL4 = "Select Distinct(pPoitem.sProdcode),pPoitem.sPono,pPo.cInvoice,pPoitem.nAmount From pPoItem,pPo,pInvoice Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono)"&_
							" and (pPo.cInvoice = 2 and pInvoice.dDate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&")"
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL4 = "Select Distinct(pPoitem.sProdcode),pPoitem.sPono,pPo.cInvoice,pPoitem.nAmount From pPoItem,pPo,pInvoice Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono)"&_
							" and (pPo.cInvoice = 2 and pInvoice.dDate >= "&To_Date(FormDate)&")"
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL4 = "Select Distinct(pPoitem.sProdcode),pPoitem.sPono,pPo.cInvoice,pPoitem.nAmount From pPoItem,pPo,pInvoice Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono)"&_
							" and (pPo.cInvoice = 2 and pInvoice.dDate <= "&To_Date(ToDate)&")"
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL4 = "Select Distinct(pPoitem.sProdcode),pPoitem.sPono,pPo.cInvoice,pPoitem.nAmount From pPoItem,pPo,pInvoice Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono)"&_
							" and pPo.cInvoice = 2"
	End IF

		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
'		Response.write sSQL4&"<br>"
		Set Rs4 = CreateObject("ADODB.Recordset")
		Rs4.Open sSQL4, Conn, adOpenStatic, adLockReadOnly, adCmdText
		PoCount = Rs4.RecordCount
		While Not Rs4.EOF
			nAmount = cDbl(Rs4("nAmount"))
			SumArr(sCount,5) = cDbl(nAmount) + cDbl(LastnAmount)
'			Response.write " Poitem = "&SumArr(sCount,5)
'			Response.write " nAmount = "&nAmount
'			Response.write " LastnAmount = "&LastnAmount&"<br>"
			LastnAmount = SumArr(sCount,5)

		Rs4.MoveNext
		Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
	Next
End Function
%>

<%
Function Invoice_PoItem(sPono)
	LastnAmount = 0
	TotalAmount = 0
	nAmount = 0
		sSQL5 = "Select * From pPoItem,pInvoice Where "&_
						" (pPoitem.sPono='"&sPono&"' and pPoitem.sPono = pInvoice.sPono)"
'		Response.write sSQL5&"<br>"
		Set Rs5 = CreateObject("ADODB.Recordset")
		Rs5.Open sSQL5, Conn, adOpenStatic, adLockReadOnly, adCmdText
			While Not Rs5.EOF
				nAmount = cDbl(Rs5("nAmount"))
				SumArr(sCount,5) = cDbl(nAmount) + cDbl(LastnAmount)
'				Response.write " Poitem = "&SumArr(sCount,5)
'				Response.write " nAmount = "&nAmount
'				Response.write " LastnAmount = "&LastnAmount&"<br>"
				LastnAmount = SumArr(sCount,5)
		Rs5.MoveNext
		Wend
	LastnAmount = 0
	TotalAmount = 0
	nAmount = 0

End Function
%>

<%
Function Get_Tax()
	For sCount = 1 to Counter

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL5 = "Select pPoitem.sPono,pPo.cTax,pPoitem.sProdcode,pPoitem.nAmount From pPoItem,pPo,pTax Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono) "&_
							" and (pPo.cTax = 2 and pTax.dDate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&")"
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL5 = "Select pPoitem.sPono,pPo.cTax,pPoitem.sProdcode,pPoitem.nAmount From pPoItem,pPo,pTax Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono) "&_
							" and (pPo.cTax = 2 and pTax.dDate >= "&To_Date(FormDate)&")"
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL5 = "Select pPoitem.sPono,pPo.cTax,pPoitem.sProdcode,pPoitem.nAmount From pPoItem,pPo,pTax Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono) "&_
							" and (pPo.cTax = 2 and pTax.dDate <= "&To_Date(ToDate)&")"
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL5 = "Select pPoitem.sPono,pPo.cTax,pPoitem.sProdcode,pPoitem.nAmount From pPoItem,pPo,pTax Where "&_
							" (pPoitem.sProdcode = '"&SumArr(sCount,11)&"' and  pPoitem.sPono=pPo.sPono) "&_
							" and pPo.cTax = 2"
		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF
	
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
'		Response.write sSQL5&"<br>"
		Set Rs5 = CreateObject("ADODB.Recordset")
		Rs5.Open sSQL5, Conn, adOpenStatic, adLockReadOnly, adCmdText
		PoCount = Rs5.RecordCount
		While Not Rs5.EOF
			nAmount = cDbl(Rs5("nAmount"))
			SumArr(sCount,6) = cDbl(nAmount) + cDbl(LastnAmount)
'			Response.write " Poitem = "&SumArr(sCount,6)
'			Response.write " nAmount = "&nAmount
'			Response.write " LastnAmount = "&LastnAmount&"<br>"
			LastnAmount = SumArr(sCount,6)

		Rs5.MoveNext
		Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
	Next
End Function
%>

<%
Function Tax_PoItem(sPono)
	LastnAmount = 0
	TotalAmount = 0
	nAmount = 0
		sSQL6 = "Select * From pPoItem,pTax Where "&_
						" (pPoitem.sPono='"&sPono&"' and pPoitem.sPono = pTax.sPono)"
'		Response.write sSQL6&"<br>"
		Set Rs6 = CreateObject("ADODB.Recordset")
		Rs6.Open sSQL6, Conn, adOpenStatic, adLockReadOnly, adCmdText
			While Not Rs6.EOF
				nAmount = cDbl(Rs6("nAmount"))
				SumArr(sCount,6) = cDbl(nAmount) + cDbl(LastnAmount)
'				Response.write " Poitem = "&SumArr(sCount,6)
'				Response.write " nAmount = "&nAmount
'				Response.write " LastnAmount = "&LastnAmount&"<br>"
				LastnAmount = SumArr(sCount,6)
		Rs6.MoveNext
		Wend
	LastnAmount = 0
	TotalAmount = 0
	nAmount = 0

End Function
%>

<%
Function Get_Receive()
	For sCount = 1 to Counter 

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL2 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPr.cPo From pPRItem,pPr,pReceive Where "&_
						" pPritem.sProdcode = '"&SumArr(sCount,11)&"' "&_
						" and pReceive.sPrno = pPr.sPrno "&_
						" and (pPr.cPo = 3 and pReceive.dRecdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&")"
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL2 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPr.cPo From pPRItem,pPr,pReceive Where "&_
						" pPritem.sProdcode = '"&SumArr(sCount,11)&"' "&_
						" and pReceive.sPrno = pPr.sPrno "&_
						" and (pPr.cPo = 3 and pReceive.dRecdate >= "&To_Date(FormDate)&")"
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL2 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPr.cPo From pPRItem,pPr,pReceive Where "&_
						" pPritem.sProdcode = '"&SumArr(sCount,11)&"' "&_
						" and pReceive.sPrno = pPr.sPrno "&_
						" and (pPr.cPo = 3 and pReceive.dRecdate <= "&To_Date(ToDate)&")"
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL2 = "Select Distinct(pPritem.sProdcode),pPritem.sPrno,pPritem.nAmount,pPr.cPo From pPRItem,pPr,pReceive Where "&_
						" pPritem.sProdcode = '"&SumArr(sCount,11)&"'"&_
						" and pReceive.sPrno = pPr.sPrno "&_
						" and pPr.cPo = 3"
		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

		
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
'		Response.write sSQL2&"<br>"
		Set Rs2 = CreateObject("ADODB.Recordset")
		Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
		While Not  Rs2.EOF
			sSQL3 = "Select * from pReceive"
			Set Rs3 = CreateObject("ADODB.Recordset")
			Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
			While Not  Rs3.EOF
				if Rs3("sPrno") = Rs2("sPrno") Then		
					nAmount = cDbl(Rs2("nAmount"))
					SumArr(sCount,4) = cDbl(nAmount) + cDbl(LastnAmount)
					LastnAmount = SumArr(sCount,4)
'					Response.write " Prno = "&Rs2("sPrno")
		'			Response.write " Pritem = "&SumArr(sCount,4)
		'			Response.write " nAmount = "&nAmount
		'			Response.write " LastnAmount = "&LastnAmount&"<br>"
				end if
			Rs3.MoveNext
			Wend
		Rs2.MoveNext
		Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
	Next
End Function
%>

<%
Function Receive_PRItem(sPrno)
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
			sSQL8 = "Select * From pPritem,pReceive,pPr Where "&_
							" (pPr.cPo = 3 and pPritem.sPrno = '"&sPrno&"') "&_
							" and pPritem.sPrno = pReceive.sPrno "
			Response.write sSQL8&"<BR>"
			Set Rs8 = CreateObject("ADODB.Recordset")	
			Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs8.RecordCount
			Response.write " nItemCount = "&nItemCount 
			While Not Rs8.EOF
				nAmount = cDbl(Rs8("nAmount"))
				SumArr(sCount,4) = cDbl(nAmount) + cDbl(LastnAmount)
				LastnAmount = SumArr(sCount,4)
				Response.write " Pritem = "&SumArr(sCount,4)
				Response.write " nAmount = "&nAmount
				Response.write " LastnAmount = "&LastnAmount&"<br>"
			Rs8.MoveNext
			Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0

End Function
%>

<%
Function SumTotalPrice(Group)
	TotalPrice = 0
	LastItem = 0
	For nCount = 1 to CattypeCount
		For sCount = 1 to Counter
			if CattypeArr(nCount) = SumArr(sCount,9) Then
				TotalPrice = SumArr(sCount,Group) + TotalPrice
'				if SumArr(sCount,0) = LastItem Then
'				else TotalPrice = SumArr(sCount,Group) + TotalPrice
'				end if
'				LastItem = SumArr(sCount,0)
			end if
		Next
	Next

	Response.write FormatNumber(TotalPrice, 3)
	TotalPrice = 0
End Function
%>

<%
Function SumGrandTotalPrice(Group)
	TotalPrice = 0
	LastItem = 0
	 	if Counter <> 0 Then 
			For sCount = 1 to Counter
				TotalPrice = SumArr(sCount,Group) + TotalPrice
			Next
		end if

	Response.write FormatNumber(TotalPrice, 3)
	TotalPrice = 0
End Function
%>
