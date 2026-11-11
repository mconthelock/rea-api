<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file ="../Connas400.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file ="../list.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo

    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,rs2,rs3
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread
	Dim svendcodefrom,svendcodeto,stname,sename,saddress1,saddress2,stelno,sfax,scontact,sposition,stnote,dadddate,sempno,sTermcode,sTermPayment,sStdCur
	Dim sSTname,sSEname,sEmail,sCity,sState,sPost,sSearchKey,sPurAddress1,sPurAddress2,sPurCity,sPurState,sPurPost,sPurContact,sPurPhone,spurEmail
	Dim img_sVendcode	,img_stname,img_sename,img_staddress,img_seaddress,img_stelno
	Dim dDate,dadddate_day,dadddate_month,dadddate_year,d_day,d_month,d_year,sbread,status
sScriptPath = Request.ServerVariables("SCRIPT_NAME")
	dDate=Date
svendcodefrom=Request("svendcodefrom")
svendcodeto=Request("svendcodeto")
'	scorpcode = Session("gbcorpcode")
'	sname = Request("ID")	
'	sUserid = Request("UID")

  ' Create a recordset and open the given event. If the Record is not found
	' set the mode to ADD
  sSQL = "SELECT * FROM AVM WHERE vendor>=" & svendcodefrom & " and vendor <="&svendcodeto 
 
  Set Rs = Server.CreateObject("ADODB.RecordSet")
    Set Rs2 = Server.CreateObject("ADODB.RecordSet")
	Set Rs3 = Server.CreateObject("ADODB.RecordSet")
'  Rs.MaxRecords = 1
  Rs.Open sSQL, connas400, adOpenStatic, adLockReadOnly, adCmdText
'response.write Rs.recordcount
 'response.write svendcodeto
IF Not Rs.EOF Then
	Do While Not Rs.EOF
		sSQL="Select * from pvender where svendcode='"& RS("vendor")&"'"
		 Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			IF Not RS2.EOF then
				Rs2("svendcode") = RS("vendor")
				Rs2("stname") = RS("vndnam")
				Rs2("sename") = RS("vndnam")
				Rs2("sEaddress") = RS("vndad1") & " " & RS("vndad2")
				Rs2("sTaddress") = RS("vndad1") & " " & RS("vndad2")
				Rs2("stelno") = RS("vphone")
'				Rs2("sfax") = sfax
				Rs2("scontact") = RS("vcon")
'				Rs2("sposition") = sposition
				RS2("sCity")=RS("vcity")
				RS2("sState")=RS("vstate")
				RS2("sPost")=RS("vpost")
				RS2("sSearchKey")=RS("vnalph")
				RS2("sPurAddr1")=RS("vpad1")
				RS2("sPurAddr2")=RS("vpad2")
				RS2("sPurCity")=RS("vpcty")
				RS2("sPurState")=RS("vpste")
				RS2("sPurPost")=RS("vppst")
				RS2("sPurContact")=RS("vpcon")
				RS2("sPurPhone")=RS("vpphn")
'				Rs2("stnote") = stnote
				Rs2("dadddate") = date
'				Rs2("sempno") = sempno
'				RS2("sSTname")=sSTname
'				RS2("sSEname")=sSEname
'				RS2("sEmail")=sEmail
'				Rs2("sPurEmail")=sPurEmail
				RS2("sTermCode")=RS("vterms")
				RS2("sStdCur")=RS("vCurr")
				RS2("sEmpno")="AS400"
				RS2.Update
			else
				  Rs3.Open "pvender", conn, adOpenKeySet, adLockPessimistic, adCmdTable
				  Rs3.AddNew
				Rs3("svendcode") = RS("vendor")
				Rs3("stname") = RS("vndnam")
				Rs3("sename") = RS("vndnam")
				Rs3("sEaddress") = RS("vndad1") & " " & RS("vndad2")
				Rs3("sTaddress") = RS("vndad1") & " " & RS("vndad2")
				Rs3("stelno") = RS("vphone")
'				Rs3("sfax") = sfax
				Rs3("scontact") = RS("vcon")
'				Rs3("sposition") = sposition
				RS3("sCity")=RS("vcity")
				RS3("sState")=RS("vstate")
				RS3("sPost")=RS("vpost")
				RS3("sSearchKey")=RS("vnalph")
				RS3("sPurAddr1")=RS("vpad1")
				RS3("sPurAddr2")=RS("vpad2")
				RS3("sPurCity")=RS("vpcty")
				RS3("sPurState")=RS("vpste")
				RS3("sPurPost")=RS("vppst")
				RS3("sPurContact")=RS("vpcon")
				RS3("sPurPhone")=RS("vpphn")
'				Rs3("stnote") = stnote
				Rs3("dadddate") = date
'				Rs3("sempno") = sempno
'				RS3("sSTname")=sSTname
'				RS3("sSEname")=sSEname
'				RS3("sEmail")=sEmail
'				Rs3("sPurEmail")=sPurEmail
				RS3("sTermCode")=RS("vterms")
				RS3("sStdCur")=RS("vCurr")
				RS3("sEmpno")="AS400"
				RS3.Update
				RS3.Close

			End IF
			RS2.Close
		RS.MoveNext
		Loop
		Response.write "Update Complete..."
Else
	response.write "No Data"
End IF

			Set RS2=nothing
			Set RS3=Nothing
			
%>