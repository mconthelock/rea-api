<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	Dim sprno,sexptype,dprdate,srequester,skey,cprtype,nstep,capprove
	Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
	Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond
	Dim Item,Prno,CashReq,sEmpno,chksEmpno,sDelivery,nitem
	sExptype=Request("sExptype")
	CashReq=Request("CashReq")
	sDelivery=Request("sDelivery")
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	sprno=Request("sNewPrno")
  	  Set Rs = Server.CreateObject("ADODB.RecordSet")	  
  	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		sSQL="Select * from ppritem where sprno='"&Request("sPrno")&"'"
		Rs2.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText	  
  		  Rs.Open "ppritem", conn, adOpenKeySet, adLockPessimistic, adCmdTable
	IF Not Rs2.EOF then
		nitem=chkitem(Request("sNewPrno"))
		  Do While Not RS2.EOF 
		  IF chkProduct(RS2("sProdcode"))="N" then
		  	nItem=cint(nItem)+1
			  Rs.AddNew
				Rs("sPrno") = sPrno
				Rs("nItem") = nItem
				Rs("sProdcode") = RS2("sProdcode")
				Rs("sProdname") = RS2("sProdname")
				Rs("sVendcode") =RS2("sVendcode")
				Rs("sProddesc") =RS2("sProddesc")
				Rs("nQuality") =RS2("nQuality")
				Rs("nPrice") =RS2("nPrice")
				Rs("nAmount") =RS2("nAmount")
				Rs("sCurcode") =RS2("sCurcode")
'				Rs("sUseFor") =RS2("sUseFor")
				Rs("sQuatation") =RS2("sQuatation")
				Rs("sUnitcode") =RS2("sUnitcode")
				Rs("sComment") =RS2("sComment")
				Rs("cStatus") ="1"
				RS("sImagespec")=RS2("sImagespec")
				RS.Update
			End IF
			 RS2.MoveNext
			 Loop
		 End IF
	  response.redirect "showprdesc.asp?sprno=" & sprno  

Function chkProduct(Product)
	Dim RS2,sSQL2,chkp

	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select sProdcode  from ppritem where  sPrno='" & Request("sNewPrno") & "' and sProdcode='" & Product &"'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF  Not Rs2.EOF then
			chkProduct="Y"
		Else
			chkProduct="N"
		End if

End Function

Function chkItem(prno)
	Dim RS2,sSQL2,chkp

	Set Rs2 = Server.CreateObject("ADODB.RecordSet")       	
	sSQL2="Select  Max(nItem)  mItem from ppritem where  sPrno='" & prno & "'"
	  Rs2.Open sSQL2, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF  Not Rs2.EOF then
			IF RS2("mitem")>"0" then
				chkItem=RS2("mitem")
			Else
				chkItem=0
			End IF
		End if
End Function
%>

