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
		Dim ExpArr(),SumArr()
		Dim Color
		Dim nItem,nItemCount,nAmount,LastnAmount,TotalAmount
		Dim ExpCount,PoCount,PrCount,sExptype
		Dim Counter,nCount,sCount,TotalPrice,sCount2,LastItem
		Dim FormDate,ToDate
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

'	IF IsDate(FormDate) and IsDate(ToDate)  Then
'		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' and dPrdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&" Order by dPrdate"
'		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
'	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
'		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' and dPrdate >= "&To_Date(FormDate)&" Order by dPrdate"
'		Message = "Data Between Begin to "&FormatDateTime(FormDate,1)
'		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
'	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
'		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' and dPrdate <= "&To_Date(ToDate)&" Order by dPrdate"
'		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
'	Else
'		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' Order by dPrdate"
'		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
'	End IF

		sSQL = "Select * From pPr"
'		Response.write sSQL
		Set Rs = CreateObject("ADODB.Recordset")
		Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
		PRCount = Rs.RecordCount

	if PRCount = 0 Then
	else
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@
					'@@@@@@@@@         Comment       @@@@@@@@@
					'	SumArr(*,0) = P/R No.
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


		sSQL6 = "Select * From pExpenstype Order by sExptype"
		Set Rs6 = CreateObject("ADODB.Recordset")
		Rs6.Open sSQL6, Conn, adOpenStatic, adLockReadOnly, adCmdText
		ExpCount = Rs6.RecordCount
		ReDim SumArr(ExpCount,20)
		While Not Rs6.EOF
		nCount = nCount + 1
			SumArr(nCount,2) = Rs6("sExptype")
'			Response.write " Exptype Code = "&SumArr(nCount,2)&"<BR>"
		Rs6.MoveNext
		Wend

		Counter = ExpCount
		nCount = ExpCount
'		Response.write " Counter = "&Counter
		Get_PRItem()
		Get_PoItem()	
		Get_Receive()
		Get_Invoice()
		Get_Tax()	
end if
%>

<HTML>
<HEAD>
<TITLE>Summary By Total</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>

<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
  <TR>
    <TD WIDTH="2%">&nbsp;</TD>
    <TD WIDTH="96%">&nbsp;</TD>
    <TD WIDTH="2%">&nbsp;</TD>
  </TR>
  <TR>
    <TD WIDTH="2%">&nbsp;</TD>
    <TD WIDTH="96%" VALIGN="TOP"> 
      <TABLE WIDTH="100%" BORDER="0" CELLSPACING="1" CELLPADDING="0">
        <TR> 
          <TD HEIGHT="11" COLSPAN="8"> 
            <TABLE WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
              <TR BGCOLOR="#3399CC"> 
                <TD height="20"> 
                  <DIV ALIGN="Left"><B><FONT COLOR="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;Summary 
                    By TOTAL : <%=Message%></FONT></B> </DIV>
                </TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
        <TR> 
          <TD WIDTH="16%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1">Expense 
              Type Code</FONT></B></DIV>
          </TD>
          <TD WIDTH="19%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="1">Expense 
              Type</FONT></B></DIV>
          </TD>
          <TD WIDTH="16%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><FONT SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif"><B>Currency</B></FONT></DIV>
          </TD>
          <TD WIDTH="9%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><FONT SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif"><B>P/R 
              AMT</B></FONT></DIV>
          </TD>
          <TD WIDTH="9%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><FONT SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif"><B>P/O 
              AMT</B></FONT></DIV>
          </TD>
          <TD WIDTH="9%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><FONT SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif"><B>Receive 
              AMT</B></FONT></DIV>
          </TD>
          <TD WIDTH="9%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><FONT SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif"><B>INV 
              AMT</B></FONT></DIV>
          </TD>
          <TD WIDTH="13%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><FONT SIZE="1" FACE="Verdana, Arial, Helvetica, sans-serif"><B>TAX 
              INV AMT</B></FONT></DIV>
          </TD>
        </TR>
        <% 
		LastItem = 0
'		For sCount = 1 to Counter
		For sCount = 1 to ExpCount
			if SumArr(sCount,2) = sCount or SumArr(sCount,2) = "0"&sCount Then
				if SumArr(sCount,8) <> "" Then
%> 
        <TR> 
          <TD BGCOLOR='<%=Color%>' WIDTH="16%" ALIGN="CENTER"><%=SumArr(sCount,2)%></TD>
          <TD bgColor='<%=Color%>' WIDTH="19%" ALIGN="CENTER"><%pExptype_Desc(SumArr(sCount,2))%></TD>
          <TD bgColor='<%=Color%>' WIDTH="16%"> 
            <DIV ALIGN="CENTER"><%Show_Currency(SumArr(sCount,3))%> </DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="9%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,8), 3)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="9%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,7), 3)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="9%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,4), 3)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="9%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,5), 3)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="13%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,6), 3)%> &nbsp;&nbsp;</DIV>
          </TD>
        </TR>
        <%
			if Color="#99CCFF" then
				Color="#FFFFFF"
			else
				Color="#99CCFF"
			end if

%> <%
			end if
			end if
		LastItem = SumArr(sCount,0)
		Next
%> 
        <TR> 
          <TD WIDTH="16%">&nbsp;</TD>
          <TD WIDTH="19%">&nbsp;</TD>
          <TD COLSPAN="6" VALIGN="BOTTOM"> 
            <HR SIZE="1">
          </TD>
        </TR>
        <TR> 
          <TD ROWSPAN="2" VALIGN="BOTTOM" COLSPAN="2"> 
            <DIV ALIGN="CENTER"><B>GRAND TOTAL</B></DIV>
          </TD>
          <TD WIDTH="16%">&nbsp;</TD>
          <TD WIDTH="9%">&nbsp;</TD>
          <TD WIDTH="9%">&nbsp;</TD>
          <TD WIDTH="9%">&nbsp;</TD>
          <TD WIDTH="9%">&nbsp;</TD>
          <TD WIDTH="13%">&nbsp;</TD>
        </TR>
        <TR> 
          <TD WIDTH="16%" bgcolor="#99CCFF" height="10">&nbsp;</TD>
          <TD WIDTH="9%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumGrandTotalPrice(8)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="9%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumGrandTotalPrice(7)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="9%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumGrandTotalPrice(4)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="9%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumGrandTotalPrice(5)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="13%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumGrandTotalPrice(6)%> &nbsp;&nbsp;</DIV>
          </TD>
        </TR>
      </TABLE>
      <BR>
      <HR SIZE="1">
    </TD>
    <TD WIDTH="2%">&nbsp;</TD>
  </TR>
  <TR>
    <TD WIDTH="2%">&nbsp;</TD>
    <TD WIDTH="96%">&nbsp;</TD>
    <TD WIDTH="2%">&nbsp;</TD>
  </TR>
</TABLE>
</BODY>
</HTML>

<%
Function Get_PoItem()

		For sCount = 1 to ExpCount 
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		nCount = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL7 = "Select Distinct(pPr.sExptype),pPo.sPono,pPoitem.sProdcode,pPr.sPrno,pPoitem.nAmount "&_
						" From pPo,pPr,pPoitem,pExpensType "&_
						" Where pPr.sExptype = '"&SumArr(sCount,2)&"' and "&_
						" pPoitem.sPono = pPo.sPono and "&_
						" pPr.sPrno = pPo.sRefno "&_
						" and (pPr.cPo = 1 and pPo.dPodate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&")" 
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL7 = "Select Distinct(pPr.sExptype),pPo.sPono,pPoitem.sProdcode,pPr.sPrno,pPoitem.nAmount "&_
						" From pPo,pPr,pPoitem,pExpensType "&_
						" Where pPr.sExptype = '"&SumArr(sCount,2)&"' and "&_
						" pPr.sPrno = pPo.sRefno and "&_
						" pPoitem.sPono = pPo.sPono "&_
						" and (pPr.cPo = 1 and pPo.dPodate >= "&To_Date(FormDate)&")" 
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL7 = "Select Distinct(pPr.sExptype),pPo.sPono,pPoitem.sProdcode,pPr.sPrno,pPoitem.nAmount "&_
						" From pPo,pPr,pPoitem,pExpensType "&_
						" Where pPr.sExptype = '"&SumArr(sCount,2)&"' and "&_
						" pPr.sPrno = pPo.sRefno and "&_
						" pPoitem.sPono = pPo.sPono "&_
						" and (pPr.cPo = 1 and pPo.dPodate <= "&To_Date(ToDate)&")" 
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL7 = "Select Distinct(pPr.sExptype),pPo.sPono,pPoitem.sProdcode,pPr.sPrno,pPoitem.nAmount "&_
						" From pPo,pPr,pPoitem,pExpensType "&_
						" Where pPr.sExptype = '"&SumArr(sCount,2)&"' and "&_
						" pPr.sPrno = pPo.sRefno and "&_
						" pPoitem.sPono = pPo.sPono "&_
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
		For sCount = 1 to ExpCount
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
'		Response.write " Counter = "&sCount&"<BR>"
'		Response.write " Exptype Code = "&SumArr(sCount,2)&"<BR>"
	IF IsDate(FormDate) and IsDate(ToDate)  Then
			sSQL8 = "Select Distinct(pExpenstype.sExptype),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,pPr,pExpenstype Where "&_
							" pExpenstype.sExptype = '"&SumArr(sCount,2)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" pExpensType.sExptype = pPr.sExptype and "&_
							" pPr.dPrdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&" and "&_
							" (pPritem.cStatus = 1 or pPritem.cStatus = 2) "&_
							" Order by pPr.dPrdate"
			Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
			sSQL8 = "Select Distinct(pExpenstype.sExptype),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,pPr,pExpenstype Where "&_
							" pExpenstype.sExptype = '"&SumArr(sCount,2)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" pExpensType.sExptype = pPr.sExptype and "&_
							" pPr.dPrdate >= "&To_Date(FormDate)&" and "&_
							" (pPritem.cStatus = 1 or pPritem.cStatus = 2) "&_
							" Order by pPr.dPrdate"
			Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
			sSQL8 = "Select Distinct(pExpenstype.sExptype),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,pPr,pExpenstype Where "&_
							" pExpenstype.sExptype = '"&SumArr(sCount,2)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" pExpensType.sExptype = pPr.sExptype and "&_
							" pPr.dPrdate <= "&To_Date(ToDate)&" and "&_
							" (pPritem.cStatus = 1 or pPritem.cStatus = 2) "&_
							" Order by pPr.dPrdate"
			Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
			sSQL8 = "Select Distinct(pExpenstype.sExptype),pPritem.sPrno,pPritem.nAmount,pPritem.sCurcode,pPritem.sProdname,pPritem.nQuality,pPritem.cStatus,pPr.dPrdate "&_
							" From pPRItem,pPr,pExpenstype Where "&_
							" pExpenstype.sExptype = '"&SumArr(sCount,2)&"' and "&_
							" pPritem.sPrno = pPr.sPrno and "&_
							" pExpensType.sExptype = pPr.sExptype and "&_
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
		sSQL4 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"' "&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
							" and (pPo.cInvoice = 2 and pInvoice.dDate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&")"
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL4 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"'"&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
							" and (pPo.cInvoice = 2 and pInvoice.dDate >= "&To_Date(FormDate)&")"
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL4 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"'"&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
							" and (pPo.cInvoice = 2 and pInvoice.dDate <= "&To_Date(ToDate)&")"
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL4 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"'"&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
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
		sSQL5 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"'"&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
							" and (pPo.cTax = 2 and pTax.dDate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&")"
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL5 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"'"&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
							" and (pPo.cTax = 2 and pTax.dDate >= "&To_Date(FormDate)&")"
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL5 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"'"&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
							" and (pPo.cTax = 2 and pTax.dDate <= "&To_Date(ToDate)&")"
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL5 = "Select Distinct(pExpenstype.sExptype),pInvoice.sInvno,pPo.sPono,pPoitem.nAmount from pInvoice,pPr,pPo,pPoitem,pExpenstype "&_
							" Where pPr.sExptype = '"&SumArr(sCount,2)&"'"&_
							" and pExpenstype.sExptype = pPr.sExptype "&_
							" and pInvoice.sPono = pPo.sPono "&_
							" and pPoitem.sPono = pPo.sPono "&_
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
		sSQL2 = "Select Distinct(pExpenstype.sExptype),pReceive.sRecno,pPr.sPrno,pPritem.nAmount From pReceive,pPr,pExpenstype,pPritem "&_
							" where pPr.sExptype = pExpenstype.sExptype "&_
							" and pPr.sExptype = '"&SumArr(sCount,2)&"' "&_
							" and pPr.sPrno = pReceive.sPrno "&_
							" and pPritem.sPrno = pPr.sPrno "&_
							" and (pPr.cPo = 3 and pReceive.dRecdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&")"
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL2 = "Select Distinct(pExpenstype.sExptype),pReceive.sRecno,pPr.sPrno,pPritem.nAmount from preceive,pPr,pExpenstype,pPritem "&_
							" where pPr.sExptype = pExpenstype.sExptype "&_
							" and pPr.sExptype = '"&SumArr(sCount,2)&"' "&_
							" and pPr.sPrno = pReceive.sPrno "&_
							" and ppritem.sPrno = pPr.sPrno "&_
							" and (pPr.cPo = 3 and pReceive.dRecdate >= "&To_Date(FormDate)&")"
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL2 = "Select Distinct(pExpenstype.sExptype),pReceive.sRecno,pPr.sPrno,pPritem.nAmount from preceive,pPr,pExpenstype,pPritem "&_
							" where pPr.sExptype = pExpenstype.sExptype "&_
							" and pPr.sExptype = '"&SumArr(sCount,2)&"' "&_
							" and pPr.sPrno = pReceive.sPrno "&_
							" and ppritem.sPrno = pPr.sPrno "&_
							" and (pPr.cPo = 3 and pReceive.dRecdate <= "&To_Date(ToDate)&")"
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL2 = "Select Distinct(pExpenstype.sExptype),pReceive.sRecno,pPr.sPrno,pPritem.nAmount from preceive,pPr,pExpenstype,pPritem "&_
							" where pPr.sExptype = pExpenstype.sExptype "&_
							" and pPr.sExptype = '"&SumArr(sCount,2)&"' "&_
							" and pPr.sPrno = pReceive.sPrno "&_
							" and ppritem.sPrno = pPr.sPrno "&_
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
'					Response.write " Pritem = "&SumArr(sCount,4)
'					Response.write " nAmount = "&nAmount
'					Response.write " LastnAmount = "&LastnAmount&"<br>"
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
			For sCount = 1 to ExpCount
				if SumArr(sCount,Group) <> "" Then
				TotalPrice = SumArr(sCount,Group) + TotalPrice
'			Response.write " TotalPrice = "&SumArr(sCount,Group)&"<br>"
				end if 
			Next
		end if

	Response.write FormatNumber(TotalPrice, 3)
	TotalPrice = 0
End Function
%>