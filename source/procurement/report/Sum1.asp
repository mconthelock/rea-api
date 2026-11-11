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
		sExptype = Request("sExptype")
		SummaryBY = Request("SummaryBY")
		FormDate = Request("Form_Month")&"/"&Request("Form_Day")&"/"&Request("Form_Year")
		ToDate = Request("To_Month")&"/"&Request("To_Day")&"/"&Request("To_Year")

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' and dPrdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)&" Order by dPrdate"
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' and dPrdate >= "&To_Date(FormDate)&" Order by dPrdate"
'		Message = "Data Between Begin to "&FormatDateTime(FormDate,1)
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' and dPrdate <= "&To_Date(ToDate)&" Order by dPrdate"
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' Order by dPrdate"
		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'		sSQL = "Select * From pPr where sExptype = '"&sExptype&"' and dPrdate = "&To_Date(FormDate)&" Order by dPrdate"
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
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@

		sSQL6 = "Select * From pExpenstype Order by sExptype"
		Set Rs6 = CreateObject("ADODB.Recordset")
		Rs6.Open sSQL6, Conn, adOpenStatic, adLockReadOnly, adCmdText
		ExpCount = Rs6.RecordCount
			sSQL3 = "Select * From pPo"
			Set Rs3 = CreateObject("ADODB.Recordset")
			Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
			PoCount = Rs3.RecordCount
			ReDim SumArr(PoCount*PrCount,10)

		For sCount = 1 to ExpCount 
		While Not RS.EOF 
		if Rs("sExptype") = "0"&sCount Then
			nCount = nCount + 1
			sSQL3 = "Select * From pPo where sRefno='"&Rs("sPrno")&"'"
			Set Rs3 = CreateObject("ADODB.Recordset")
			Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
			PoCount = Rs3.RecordCount
			if PoCount = 0 Then
								SumArr(nCount,0) = Rs("sPrno")
								SumArr(nCount,1) = ""
								SumArr(nCount,2) = Rs("sExptype")
			else
				While Not  Rs3.EOF
					Do While Not PoCount = 0
								SumArr(nCount,0) = Rs("sPrno")
								SumArr(nCount,1) = Rs3("sPono")
								SumArr(nCount,2) = Rs("sExptype")
							PoCount = PoCount - 1
						if PoCount = 0 Then
						else nCount = nCount + 1
						end if
					Loop
				Rs3.MoveNext
				Wend	
			end if
		else
			if Rs("sExptype") = sCount Then
				nCount = nCount + 1
				sSQL3 = "Select * From pPo where sRefno='"&Rs("sPrno")&"'"
				Set Rs3 = CreateObject("ADODB.Recordset")
				Rs3.Open sSQL3, Conn, adOpenStatic, adLockReadOnly, adCmdText
				PoCount = Rs3.RecordCount
					if PoCount = 0 Then
						SumArr(nCount,0) = Rs("sPrno")
						SumArr(nCount,1) = ""
						SumArr(nCount,2) = Rs("sExptype")
					else
						While Not  Rs3.EOF
							Do While Not PoCount = 0
								SumArr(nCount,0) = Rs("sPrno")
								SumArr(nCount,1) = Rs3("sPono")
								SumArr(nCount,2) = Rs("sExptype")
							PoCount = PoCount - 1
								if PoCount = 0 Then
								else nCount = nCount + 1
								end if
						Loop
						Rs3.MoveNext
						Wend	
					end if
					Rs3.Close
			end if
		end if
		Rs.MoveNext
		Wend
		Rs.MoveFirst
		Next

		Counter = nCount
		Get_Receive()
		Get_PoItem()	
		Get_PRItem()

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
          <TD HEIGHT="11" COLSPAN="7"> 
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
          <TD WIDTH="28%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B>Expense Type</B></DIV>
          </TD>
         <TD WIDTH="12%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B>Currency</B></DIV>
          </TD>
          <TD WIDTH="12%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B>P/R AMT</B></DIV>
          </TD>
          <TD WIDTH="12%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B>P/O AMT</B></DIV>
          </TD>
          <TD WIDTH="12%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B>Receive AMT</B></DIV>
          </TD>
          <TD WIDTH="12%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B>INV AMT</B></DIV>
          </TD>
          <TD WIDTH="12%" BGCOLOR="#33CCFF"> 
            <DIV ALIGN="CENTER"><B>TAX INV AMT</B></DIV>
          </TD>
        </TR>
        <% 
		LastItem = 0
		For sCount = 1 to Counter
			if SumArr(sCount,0) <> LastItem Then
%> 
        <TR> 
          <TD bgColor='<%=Color%>' WIDTH="28%">&nbsp;&nbsp;&nbsp;<%pExptype_Desc(SumArr(sCount,2))%></TD>
          <TD bgColor='<%=Color%>' WIDTH="12%"> 
            <DIV ALIGN="CENTER"><%Show_Currency(SumArr(sCount,3))%> </DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="12%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,8),0)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="12%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,7),0)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="12%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,4),0)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="12%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,5),0)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD bgColor='<%=Color%>' WIDTH="12%"> 
            <DIV ALIGN="Right"><%=FormatNumber(SumArr(sCount,6),0)%> &nbsp;&nbsp;</DIV>
          </TD>
        </TR>
<%
			if Color="#99CCFF" then
				Color="#FFFFFF"
			else
				Color="#99CCFF"
			end if

%>
<%
		end if
		LastItem = SumArr(sCount,0)
		Next
%> 
        <TR> 
          <TD WIDTH="28%">&nbsp;</TD>
          <TD COLSPAN="6" VALIGN="BOTTOM"> 
            <HR SIZE="1">
          </TD>
        </TR>
        <TR> 
          <TD ROWSPAN="2" VALIGN="BOTTOM" WIDTH="28%"> 
            <DIV ALIGN="CENTER"><B>GRAND TOTAL</B></DIV>
          </TD>
          <TD WIDTH="12%">&nbsp;</TD>
          <TD WIDTH="12%">&nbsp;</TD>
          <TD WIDTH="12%">&nbsp;</TD>
          <TD WIDTH="12%">&nbsp;</TD>
          <TD WIDTH="12%">&nbsp;</TD>
          <TD WIDTH="12%">&nbsp;</TD>
        </TR>
        <TR> 
          <TD WIDTH="12%" bgcolor="#99CCFF" height="10">&nbsp;</TD>
          <TD WIDTH="12%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumTotalPrice(8)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="12%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumTotalPrice(7)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="12%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumTotalPrice(4)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="12%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumTotalPrice(5)%> &nbsp;&nbsp;</DIV>
          </TD>
          <TD WIDTH="12%" bgcolor="#99CCFF" height="10"> 
            <DIV ALIGN="Right"><%=SumTotalPrice(6)%> &nbsp;&nbsp;</DIV>
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
Function Get_Receive()
	For sCount = 1 to Counter 
		sSQL2 = "Select * From pReceive where sPrno = '"&SumArr(sCount,0)&"'"
		Set Rs2 = CreateObject("ADODB.Recordset")
		Rs2.Open sSQL2, Conn, adOpenStatic, adLockReadOnly, adCmdText
		
		While Not  Rs2.EOF
			Receive_PRItem(Rs2("sPrno"))
		Rs2.MoveNext
		Wend
	Next
End Function
%>

<%
Function Get_PoItem()

		For sCount = 0 to Counter 
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		sSQL7 = "Select ppr.sprno,ppo.srefno,ppo.spono,ppoitem.nAmount from ppr,ppo,ppoitem "&_
						"where (ppo.spono=ppoitem.spono) and (ppo.srefno=ppr.sprno)"
			Set Rs7 = CreateObject("ADODB.Recordset")	
			Rs7.Open sSQL7, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs7.RecordCount

			While Not Rs7.EOF
			if Rs7("sPrno") = SumArr(sCount,0) Then
				nAmount = cDbl(Rs7("nAmount"))
				TotalAmount = cDbl(nAmount)
				SumArr(sCount,7) = cDbl(TotalAmount) + cDbl(LastnAmount)
'				Response.write " Pono = "&SumArr(sCount,1)
'				Response.write " Poitem = "&SumArr(sCount,7)
'				Response.write " nAmount = "&nAmount
'				Response.write " nItemCount = "&nItemCount
'				Response.write " TotalAmount = "&TotalAmount
'				Response.write " LastnAmount = "&LastnAmount
'				Response.write "<BR>"
				LastnAmount = cDbl(Rs7("nAmount"))
			end if 
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

		For sCount = 0 to Counter 
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
			sSQL8 = "Select sPrno,nAmount,sCurcode From pPRItem Where sPrno = '"&SumArr(sCount,0)&"' and cStatus <> 0"
			Set Rs8 = CreateObject("ADODB.Recordset")	
			Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs8.RecordCount
			While Not Rs8.EOF
				nAmount = cDbl(Rs8("nAmount"))
				TotalAmount = cDbl(nAmount)
				SumArr(sCount,8) = cDbl(TotalAmount) + cDbl(LastnAmount)
'				Response.write " Prno = "&SumArr(sCount,0)
'				Response.write " Pritem = "&SumArr(sCount,8)
'				Response.write " nAmount = "&nAmount
'				Response.write " TotalAmount = "&TotalAmount
'				Response.write " LastnAmount = "&LastnAmount
'				Response.write "<BR>"
				LastnAmount = cDbl(Rs8("nAmount"))
				SumArr(sCount,3) = Rs8("sCurcode")
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
		sSQL4 = "Select * From pInvoice Where sPono='"&SumArr(sCount,1)&"'"
		Set Rs4 = CreateObject("ADODB.Recordset")
		Rs4.Open sSQL4, Conn, adOpenStatic, adLockReadOnly, adCmdText

		While Not Rs4.EOF
			Invoice_PoItem(Rs4("sPono"))
		Rs4.MoveNext
		Wend
	Next
End Function
%>

<%
Function Get_Tax()
	For sCount = 1 to Counter
		sSQL5 = "Select * From pTax Where sPono='"&SumArr(sCount,1)&"'"
		Set Rs5 = CreateObject("ADODB.Recordset")
		Rs5.Open sSQL5, Conn, adOpenStatic, adLockReadOnly, adCmdText
		
		While Not Rs5.EOF
			Tax_PoItem(Rs5("sPono"))
		Rs5.MoveNext
		Wend
	Next
End Function
%>

<%
Function Receive_PRItem(sPrno)
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
			sSQL8 = "Select * From pPritem Where sPrno = '"&sPrno&"' and cStatus <> 0"
			Set Rs8 = CreateObject("ADODB.Recordset")	
			Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs8.RecordCount
			
			While Not Rs8.EOF
				nAmount = cDbl(Rs8("nAmount"))
				TotalAmount = cDbl(nAmount)
				SumArr(sCount,4) = cDbl(TotalAmount) + cDbl(LastnAmount)
				LastnAmount = cDbl(Rs8("nAmount"))
			Rs8.MoveNext
			Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0

End Function
%>

<%
Function Invoice_PoItem(sPono)
	LastnAmount = 0
	TotalAmount = 0
	nAmount = 0
		sSQL5 = "Select * From pPoItem Where sPono='"&sPono&"'"
		Set Rs5 = CreateObject("ADODB.Recordset")
		Rs5.Open sSQL5, Conn, adOpenStatic, adLockReadOnly, adCmdText
			While Not Rs5.EOF
				nAmount = cDbl(Rs5("nAmount"))
				TotalAmount = cDbl(nAmount)
				SumArr(sCount,5) = cDbl(TotalAmount) + cDbl(LastnAmount)
				LastnAmount = cDbl(Rs5("nAmount"))
		Rs5.MoveNext
		Wend
	LastnAmount = 0
	TotalAmount = 0
	nAmount = 0

End Function
%>

<%
Function Tax_PoItem(sPono)
	LastnAmount = 0
	TotalAmount = 0
	nAmount = 0
		sSQL6 = "Select * From pPoitem Where sPono='"&sPono&"'"
		Set Rs6 = CreateObject("ADODB.Recordset")
		Rs6.Open sSQL6, Conn, adOpenStatic, adLockReadOnly, adCmdText

		While Not Rs6.EOF
			nAmount = cDbl(Rs6("nAmount"))
			TotalAmount = cDbl(nAmount)
			SumArr(sCount,6) = cDbl(TotalAmount) + cDbl(LastnAmount)
			LastnAmount = cDbl(Rs6("nAmount"))
		Rs6.MoveNext
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
	For sCount = 1 to Counter
		if SumArr(sCount,0) = LastItem Then
		else TotalPrice = SumArr(sCount,Group) + TotalPrice
		end if
	LastItem = SumArr(sCount,0)
	Next 
	Response.write FormatNumber(TotalPrice,0)
	TotalPrice = 0
End Function
%>