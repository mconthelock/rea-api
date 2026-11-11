<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="scm.inc" -->
<%
'on error resume next
Dim RS,sSQL,sSQL2,RS2,RS3,RS4,RS5,RS6,RS7,RS1,RS8
Dim sPrno,sRequester,sKey,svend,sEname,sEaddress,sDelivery
Dim img_sprno,img_step,img_dprdate_year,sbread,d_day,d_month,d_year,dprdate_day,dprdate_month,dprdate_year
Dim dDate,dDay,dMonth,dYear,dHour,dMinute,dSecond,cMonth,Pono,sPono(),svendcode(),sRefno(),nCount,i,sCurcode()
sPrno=Request("sPrno")
Set Rs=Server.CreateObject("ADODB.Recordset")
Set Rs1=Server.CreateObject("ADODB.Recordset")
Set RS2=Server.CreateObject("ADODB.Recordset")
Set RS3=Server.CreateObject("ADODB.Recordset")
'Set Rs8=Server.CreateObject("ADODB.Recordset")
sSQL = "select * from amec.ppo where srefno = '"&sPrno&"' "
Rs1.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
IF NOT Rs1.EOF then
	Response.write "Duplicate PO."
	response.end
end if
Rs1.Close

sSQL="Select * from ppr where sprno='" & sPrno & "'"
 Rs1.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
IF Not RS1.EOF then
	Do While Not RS1.EOF
		sRequester=RS1("sRequester")
		sKey=RS1("sKey")
		sDelivery=Rs1("sDelivery")
	RS1.MoveNext
	Loop
Rs1.Close
sSQL="Select distinct svendcode from ppritem Where cStatus='1'" &_
	"and sPrno='" & sPrno & "'"
 Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
If Not RS.EOF then

	Rs3.Open "ppo", conn, adOpenKeySet, adLockPessimistic, adCmdTable
	Do While Not RS.EOF
		svend=RS("sVendcode")
			sSQL="Select distinct sCurcode from pPritem where sVendcode='" & svend & "'" &_
			" and  cStatus='1'" &_
			" and sPrno='" & sPrno & "'"
 			Rs2.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText		
			IF Not RS2.EOF then
				Do While Not Rs2.EOF
		  		    PoNo = CreatePONO	
					RS3.AddNew
					RS3("sPono")=Pono
					RS3("dPodate")=Date
					RS3("sRefno")=sPrno
					RS3("sCurrency")=RS2("sCurcode")
					RS3("sRequester")=sRequester
					RS3("sKey")=sKey
					RS3("sDelivery")=sDelivery

					if request("dDelivery") <> "" and request("mDelivery") <> "" and request("yDelivery") <> "" then
						'RS3("dDelivery") = request("mDelivery") & "/" & request("dDelivery") & "/" & request("yDelivery")
						'change format from mm/dd/yyyy to dd/mm/yyyy	(09/09/2009') by Kallayanee
						RS3("dDelivery") =  request("dDelivery") & "/" & request("mDelivery") & "/" & request("yDelivery")
					end if

					RS3("sVendcode")=RS("sVendcode")
					Get_Vender(RS("sVendcode"))
					RS3("sEname")=sEname
					Rs3("sEaddress")=sEaddress
					RS3("sByer")="0"
					RS3("cstatus")="0"
					RS3("cReceive")="0"
					'RS3("sShipto")="MITSUBISHI ELEVATOR ASIA CO.,LTD."
					RS3("sShipto")=getDefaultShipTo(sRequester, conn)
					RS3("sPayterm")="60 Days After Delivery"
					RS3.Update
				RS2.MoveNext
				Loop
			End If
			RS2.Close
	RS.MoveNext
	Loop
End if ' End RS
End if ' End RS1

Set RS4=Server.CreateObject("ADODB.Recordset")
Set RS5=Server.CreateObject("ADODB.Recordset")
Set RS6=Server.CreateObject("ADODB.Recordset")
sSQL="Select * from ppo where sRefno='" & sprno & "'"
 Rs4.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
 Rs5.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
 nCount=Rs5.RecordCount
 Rs5.Close

IF Not RS4.EOF then
		reDim sPono(nCount)
		reDim sVendcode(nCount)
		reDim sRefno(nCount) 
		reDim sCurcode(nCount)
		i=0
	Do While Not RS4.EOF
		sPono(i)=RS4("sPono")
		sVendcode(i)=RS4("sVendcode")
		sRefno(i)=RS4("sRefno")
		sCurcode(i)=RS4("sCurrency")
	RS4.MoveNext
	i=i+1
	Loop
End if
		Rs6.Open "ppoitem", conn, adOpenKeySet, adLockPessimistic, adCmdTable
For i=0 to nCount-1
			sSQL="Select * from ppritem where cStatus='1'" &_
			" and sVendcode='" & sVendcode(i) & "'" &_
			" and sprno='" & sRefno(i) & "'" &_
			" and sCurcode='" & sCurcode(i) & "'"
			 Rs5.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
			IF Not RS5.EOF then
					Do While Not RS5.EOF
						RS6.AddNew
						RS6("sPono")=sPono(i)
						RS6("sProdcode")=RS5("sProdcode")
						RS6("sProdname")=RS5("sProdname")
						RS6("sProddesc")=RS5("sProddesc")
						RS6("nQuality")=RS5("nQuality")
						RS6("nPrice")=RS5("nPrice")
						RS6("nAmount")=RS5("nAmount")
						RS6("sCurcode")=RS5("sCurcode")
						RS6("sUnitcode")=RS5("sUnitcode")
						RS6("sComment")=RS5("sComment")
						RS6("cstatus")="1"
						Rs6.Update
					RS5.MoveNext
					Loop
			End if
			Rs5.Close
next
Set RS7=Server.CreateObject("ADODB.Recordset")
sSQL="Select * from ppr Where sprno='" & sPrno & "'"
 Rs7.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
If Not RS7.EOF then
	Do While Not Rs7.EOF
		RS7("cpo") = "1"
		RS7.Update
	Rs7.MoveNext
	Loop
End if
RS7.Close
response.redirect "../workflow/workflow.asp"

Sub Get_Vender(Vendcode)
Dim sSQL2,RS2,cSelect
		sSQL2="Select * from pvender where sVendcode='" & Trim(Vendcode) & "'" &_
		"Order By svendcode"
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		Rs2.Open sSQL2, conn, adOpenStatic, adLockReadOnly, adCmdText

	IF NOT Rs2.EOF  Then
		sEname=Rs2("sEname")
		sEaddress=Rs2("sEaddress")
	End IF
End Sub
%>