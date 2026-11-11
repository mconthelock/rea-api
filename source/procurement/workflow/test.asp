<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<!-- #include file="../List.asp" -->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<%
'sendMail "amecoutlook", "adminprocurement", "adminprocurement@MitsubishiElevatorAsia.co.th", "chatch@MitsubishiElevatorAsia.co.th", "Test3", "Test3" & " Sent Email"	
sendMail "amecoutlook", "adminprocurement", "adminprocurement@MitsubishiElevatorAsia.co.th", "chatch@MitsubishiElevatorAsia.co.th", "Test auto from procurement", "Test auto from procurement" & " Sent Email"	
'sendMail "amecoutlook", "adminprocurement", "adminprocurement@MitsubishiElevatorAsia.co.th", "kanittha@MitsubishiElevatorAsia.co.th", "Test", "Test" & " Sent Email"
response.write "xxxx"
response.end	
   Dim  Rs,sSQL,sSQL2,RS2
	Dim sPrno,sExptype,dPrdate,nPrcount,i
	Dim Color,noPr
	Dim Prno(),Step(),Approve(),Appdate()
	Color="#F1B3A7"
	Set Rs = Server.CreateObject("ADODB.RecordSet")

	sSQL = "SELECT  * FROM pprapp where sprno='000726122041' and sempno='97058'" 
 	Rs.Open sSQL, Conn,adOpenKeySet, adLockPessimistic, adCmdText
if Not RS.EOF then
	Do While Not Rs.EOF 
		'	RS("nStep")=4
			RS("cApprov")="1"
			RS("dappdate")=# 7/27/00 #
			RS.Update
		RS.MoveNext
	Loop
end if
response.write now+7
response.write now
response.write day(now+7) 
%>