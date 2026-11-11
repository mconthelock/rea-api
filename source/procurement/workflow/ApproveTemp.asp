<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->

<%
	' Author: Ekkapan Kongchoo
	' Version: 1.00
%>	

<%
	Dim sErrImage, Rs, sSQL,RS7,RS2,rs3,Rs5,sSQL5
	Dim sPrno, nItem, Counter
	Dim	cStatus(100)
	Dim nStep,MaxStep,NextStep
	Dim tAppTime,dAppDate, cApprov,cPrtype,LastStep,chkitem

		sPrno = Request.Form("sPrno")
		nItem = Request.Form("nItem")
		Counter = Request.Form("Count_Item")

		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		sSQL="Select nItem from ppritem where sPrno='" & sPrno &_
		"' and cstatus<>'0'" &_
		" Order By nItem"
		Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Do While Not RS2.EOF
			nItem=RS2("nItem")
	'While Not  
Counter = 0
		'nItem = Counter


		sSQL = "SELECT * FROM  ppritem WHERE sPrno= '" & sPrno &"' and nItem = " & nItem
		Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF Request("cStatus_"&nItem)="0" then
		 RS("sUsefor")="Reject By " & Session("sEmpno")
		End if
		Rs("cStatus") = Request("cStatus_"&nItem)
		Rs("sNote") = Request("sNote_"&nItem)
		Rs.Update
		Counter = Counter +1
		Rs.Close
	RS2.MoveNext
	Loop

		Set Rs = Nothing
Response.Write "Update Complete"

%>