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
		Dim SumArr(),RequesterArr(),ProductArr(),ExpensTypeArr()
		Dim Color
		Dim nItem,nItemCount,nAmount,LastnAmount,TotalAmount
		Dim Quality,LastQuality
		Dim CatIDCount,CattypeCount,sExptype
		Dim Counter,nCount,sCount,TotalPrice,sCount2,LastItem
		Dim FormDate,ToDate
		Dim PrCount,PoCount,RequesterCount,ProductCount,VendorCount,ExpCount 
		Dim SummaryBY,Message
		Dim VendorCode
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
		Quality = 0
		LastQuality = 0
%>
<%
		sExptype = Request("sExptype")
		SummaryBY = Request("SummaryBY")
		FormDate = Request("Form_Month")&"/"&Request("Form_Day")&"/"&Request("Form_Year")
		ToDate = Request("To_Month")&"/"&Request("To_Day")&"/"&Request("To_Year")
		VendorCode = Request("VendCode")
'		Response.write " VendorCode = "&VendorCode&"<br>"
%>
<%

					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@
					'@@@@@@@@@         Comment      @@@@@@@@@@@
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
					'	SumArr(*,13) = Requester ID ( sEmpno )
					'	SumArr(*,14) = Vendor Code
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@
					'@@@@@@@@@@@@@@@@@@@@@@@@@@@@

		nCount = 0
		sSQL6 = "Select * From pExpenstype Order by sExptype"
		Set Rs6 = CreateObject("ADODB.Recordset")
		Rs6.Open sSQL6, Conn, adOpenStatic, adLockReadOnly, adCmdText
		ExpCount = Rs6.RecordCount
		ReDim ExpensTypeArr(ExpCount,10) 
		While Not Rs6.EOF
		nCount = nCount + 1
			ExpensTypeArr(nCount,1) = Rs6("sExptype")
			ExpensTypeArr(nCount,2) = Rs6("sExpdesc")
		Rs6.MoveNext
		Wend

		nCount = 0
		sSQL8 = "select Distinct sVendcode,sExptype From vPpr Where sVendcode ='"&VendorCode&"'"
		Set Rs8 = CreateObject("ADODB.Recordset")
		Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText		
		VendorCount = Rs8.RecordCount 
'		Response.write " VendorCount = "&VendorCount
		ReDIM SumArr(VendorCount*ExpCount,20)
		While Not Rs8.EOF
		nCount = nCount + 1
			SumArr(nCount,2) = Rs8("sExptype")
			SumArr(nCount,14) = Rs8("sVendcode")
		Rs8.MoveNext
		Wend 

		Get_PRItem()
		Get_PoItem()
		Get_Receive()
		Get_Invoice()
		Get_Tax()	
%>

<HTML>
<HEAD>
<TITLE>Summary By Requester Section</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=">
<LINK REL="stylesheet" HREF="../styles.css">
</HEAD>

<BODY BGCOLOR="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<B><FONT FACE="Verdana, Arial, Helvetica, sans-serif" SIZE="2" COLOR="#FFFFFF"><%=ProductCount%></FONT></B>
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
                <td HEIGHT="10"> <b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;Expense 
                  By Vendor</font></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#33CCFF" width="7%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">MM/YYYY</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="9%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">VENDOR</font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="11%"> 
            <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Expense 
              Type </font></b></div>
          </td>
          <td bgcolor="#33CCFF" width="7%"> 
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
For sCount = 1 to VendorCount
'	For sCount = 1 to ProductCount
'		Response.write nCount&"<br>"
'		Response.write sCount&"<br>"
'		Response.write RequesterArr(nCount,1) &" = "
'		Response.write SumArr(sCount,13)&" "

'		if RequesterArr(nCount,1) = SumArr(sCount,13) Then
%> 
        <tr> 
          <td bgcolor='<%=Color%>' width="7%" height="10" align="center"><%'=RequesterArr(nCount,1)%> 
			</td>
          <td bgcolor='<%=Color%>' width="9%" height="10" align="center"> <%Show_Vender(SumArr(sCount,14))%> 
          </td>
          <td bgcolor='<%=Color%>' width="11%" height="10" align="center"> <%pExptype_Desc(SumArr(sCount,2))%> 
          </td>
          <td bgcolor='<%=Color%>' width="7%" height="10" align="center"> <%Show_Currency(SumArr(sCount,3))%> 
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
%> <%
'			end if
		if sCount = VendorCount Then 
%> 
        <tr> 
          <td width="7%">&nbsp;</td>
          <td width="7%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td valign="BOTTOM" colspan="6"> 
            <hr size="1">
          </td>
        </tr>
        <tr> 
          <td width="7%"> 
            <div align="CENTER"></div>
          </td>
          <td COLSPAN="2"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b> 
              TOTAL</b></font></div>
          </td>
          <td width="7%" bgcolor="#99CCFF" Align="Center"> <%Show_Currency(SumArr(sCount,3))%></td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(8)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(7)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(4)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(5)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" Align="Right"><%SumTotalPrice(6)%>&nbsp;</td>
        </tr>
        <% 
		end if
'		end if
'	Next
Next 
%> 
<!---
        <tr> 
          <td width="7%">&nbsp;</td>
          <td width="9%">&nbsp;</td>
          <td width="11%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
          <td width="10%">&nbsp;</td>
        </tr>
        <tr> 
          <td width="7%" height="16">&nbsp;</td>
          <td COLSPAN="2" height="16"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>GRAND 
              TOTAL</b></font></div>
          </td>
          <td width="7%" bgcolor="#99CCFF" height="16" ALIGN="CENTER"><%if ProductCount <> 0 Then Show_Currency(SumArr(sCount,3))%></td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if ProductCount  <> 0 Then SumGrandTotalPrice(8)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if ProductCount  <> 0 Then SumGrandTotalPrice(7)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if ProductCount  <> 0 Then SumGrandTotalPrice(4)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if ProductCount  <> 0 Then SumGrandTotalPrice(5)%>&nbsp;</td>
          <td width="10%" bgcolor="#99CCFF" height="16" ALIGN="RIGHT"><%if ProductCount  <> 0 Then SumGrandTotalPrice(6)%>&nbsp;</td>
        </tr>
        <tr> 
          <td width="7%" height="14">&nbsp;</td>
          <td width="9%" height="14">&nbsp;</td>
          <td width="11%" height="14">&nbsp;</td>
          <td width="7%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
          <td width="10%" height="14">&nbsp;</td>
        </tr>
//-->
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

		For sCount = 1 to VendorCount
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		nCount = 0
		Quality = 0 
		LastQuality = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.dPodate Between "&To_Date(FormDate)&" and "&To_Date(ToDate) 
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.dPodate >= "&To_Date(FormDate)
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono"&_
						" and pPo.dPodate <= "&To_Date(ToDate)
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "
		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'			Response.write sSQL7&"<BR>"
			Set Rs7 = CreateObject("ADODB.Recordset")	
			Rs7.Open sSQL7, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs7.RecordCount

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
		Quality = 0 
		LastQuality = 0
		Next

End Function
%>

<%
Function Get_PRItem()

		nCount = 1
		For sCount = 1 to VendorCount
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Quality = 0 
		LastQuality = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&VendorCode&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and (cStatus = 1 or cStatus = 2) "&_
							" and dPrdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)
			Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&VendorCode&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and (vPpr.cStatus = 1 or vPpr.cStatus = 2) "&_
							" and vpPr.dPrdate >= "&To_Date(FormDate)
			Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&VendorCode&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and (vPpr.cStatus = 1 or vPpr.cStatus = 2) "&_
							" and vpPr.dPrdate <= "&To_Date(ToDate)
			Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&VendorCode&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and (vPpr.cStatus = 1 or vPpr.cStatus = 2) "&_
							" and (vPpr.cStatus = 1 or vPpr.cStatus = 2) "
			Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'			Response.write sSQL8&"<BR>"
			Set Rs8 = CreateObject("ADODB.Recordset")	
			Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs8.RecordCount
			While Not Rs8.EOF

					SumArr(sCount,3) = Rs8("sCurcode")
					nAmount = cDbl(Rs8("nAmount"))
'					Quality = cDbl(Rs8("nQuality"))
'					SumArr(sCount,12) = cDbl(Quality) + cDbl(LastQuality)
					SumArr(sCount,8) = cDbl(nAmount) + cDbl(LastnAmount)
'					Response.write " Requester = "&SumArr(sCount,13)
'					Response.write " Prodcode = "&SumArr(sCount,11)
'					Response.write " Pritem = "&SumArr(sCount,8)
'					Response.write " Currency = "&SumArr(sCount,3)
'					Response.write " nAmount = "&nAmount
'					Response.write " LastnAmount_1 = "&LastnAmount
'					Response.write " sCount = "&sCount
'					Response.write "<BR>"
					LastnAmount = SumArr(sCount,8)
'					LastQuality = SumArr(sCount,12)

			Rs8.MoveNext
			Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Quality = 0 
		LastQuality = 0
		Next

End Function
%>

<%
Function Get_InVoice()

		For sCount = 1 to ProductCount 
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		nCount = 0
		Quality = 0 
		LastQuality = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sInvno <> '""' "&_
						" and vPpo.dPodate Between "&To_Date(FormDate)&" and "&To_Date(ToDate) 
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sInvno <> '""' "&_
						" and vPpo.dPodate >= "&To_Date(FormDate)
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sInvno <> '""' "&_
						" and pPo.dPodate <= "&To_Date(ToDate)
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sInvno <> '""' "
		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'			Response.write sSQL7&"<BR>"
			Set Rs7 = CreateObject("ADODB.Recordset")	
			Rs7.Open sSQL7, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs7.RecordCount

			While Not Rs7.EOF
				nAmount = cDbl(Rs7("nAmount"))
				SumArr(sCount,5) = cDbl(nAmount) + cDbl(LastnAmount)
'				Response.write " Poitem = "&SumArr(sCount,5)
'				Response.write " nAmount = "&nAmount
'				Response.write " nItemCount = "&nItemCount
'				Response.write " LastnAmount = "&LastnAmount
'				Response.write "<BR>"
				LastnAmount = SumArr(sCount,5)
			Rs7.MoveNext
			Wend

		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Quality = 0 
		LastQuality = 0
		Next

End Function
%>

<%
Function Get_Tax()

		For sCount = 1 to ProductCount 
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		nCount = 0
		Quality = 0 
		LastQuality = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sTaxno <> '""' "&_
						" and vPpo.dPodate Between "&To_Date(FormDate)&" and "&To_Date(ToDate) 
		Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sTaxno <> '""' "&_
						" and vPpo.dPodate >= "&To_Date(FormDate)
		Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sTaxno <> '""' "&_
						" and pPo.dPodate <= "&To_Date(ToDate)
		Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
		sSQL7 = " Select Distinct vPpo.sPono,vPpo.sRefno,pPoitem.sProdcode,pPoitem.nAmount,vPpr.sExptype "&_
						" From vPpo,pPoitem,vPpr "&_
						" where vPpr.sExptype = '"&SumArr(sCount,2)&"'"&_
						" and vPpo.sVendcode ='"&VendorCode&"' "&_
						" and vPpo.sPono = pPoitem.sPono "&_
						" and vPpo.sTaxno <> '""' "
		Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'			Response.write sSQL7&"<BR>"
			Set Rs7 = CreateObject("ADODB.Recordset")	
			Rs7.Open sSQL7, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs7.RecordCount

			While Not Rs7.EOF
				nAmount = cDbl(Rs7("nAmount"))
				SumArr(sCount,6) = cDbl(nAmount) + cDbl(LastnAmount)
'				Response.write " Poitem = "&SumArr(sCount,6)
'				Response.write " nAmount = "&nAmount
'				Response.write " nItemCount = "&nItemCount
'				Response.write " LastnAmount = "&LastnAmount
'				Response.write "<BR>"
				LastnAmount = SumArr(sCount,6)
			Rs7.MoveNext
			Wend

		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Quality = 0 
		LastQuality = 0
		Next

End Function
%>

<%
Function Get_Receive()

		nCount = 1
		For sCount = 1 to VendorCount
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Quality = 0 
		LastQuality = 0

	IF IsDate(FormDate) and IsDate(ToDate)  Then
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&SumArr(sCount,14)&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and cPo = 3 "&_							
							" and dPrdate Between "&To_Date(FormDate)&" and "&To_Date(ToDate)
			Message = "Data Between"&FormatDateTime(FormDate,1)&" and "&FormatDateTime(ToDate,1)
	Elseif  IsDate(FormDate) and Not IsDate(ToDate) Then
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&SumArr(sCount,14)&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and cPo = 3 "&_
							" and dPrdate >= "&To_Date(FormDate)
			Message = "Data Between "&FormatDateTime(FormDate,1)&" to  "&FormatDateTime(Now(),1)
	Elseif  Not IsDate(FormDate) and IsDate(ToDate) Then
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&SumArr(sCount,14)&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and cPo = 3 "&_
							" and dPrdate <= "&To_Date(ToDate)
			Message = "Data Between Begin to "&FormatDateTime(ToDate,1)
	Else
			sSQL8 = " Select nAmount,sCurcode "&_
							" From vPPr "&_
							" Where sVendcode = '"&SumArr(sCount,14)&"' "&_
							" and sExptype = '"&SumArr(sCount,2)&"'"&_
							" and cPo = 3 "
			Message = "Data Between Begin to "&FormatDateTime(Now(),1)
	End IF

'			Response.write sSQL8&"<BR>"
			Set Rs8 = CreateObject("ADODB.Recordset")	
			Rs8.Open sSQL8, Conn, adOpenStatic, adLockReadOnly, adCmdText
			nItemCount = Rs8.RecordCount
			While Not Rs8.EOF

					nAmount = cDbl(Rs8("nAmount"))
					SumArr(sCount,4) = cDbl(nAmount) + cDbl(LastnAmount)
					LastnAmount = SumArr(sCount,4)

			Rs8.MoveNext
			Wend
		LastnAmount = 0
		TotalAmount = 0
		nAmount = 0
		Quality = 0 
		LastQuality = 0
		Next

End Function
%>

<%
Function SumTotalPrice(Group)
	TotalPrice = 0
	LastItem = 0
'	For nCount = 1 to VendorCount
		For sCount = 1 to VendorCount
'			if RequesterArr(nCount,1) = ProductArr(sCount,1) Then
				TotalPrice = SumArr(sCount,Group) + TotalPrice
'			end if
		Next
'	Next

	Response.write FormatNumber(TotalPrice, 3)
	TotalPrice = 0
End Function
%>

<%
Function SumGrandTotalPrice(Group)
	TotalPrice = 0
	LastItem = 0
	 	if ProductCount <> 0 Then 
			For sCount = 1 to ProductCount*RequesterCount
				TotalPrice = SumArr(sCount,Group) + TotalPrice
			Next
		end if

	Response.write FormatNumber(TotalPrice, 3)
	TotalPrice = 0
End Function
%>
