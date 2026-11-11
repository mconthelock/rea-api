<% Option Explicit %>
<!-- #include file="..\adovbs.inc" -->
<!-- #include file ="..\Conn.asp" -->
<!-- #include file ="..\Connas400.asp" -->
<!-- #include file="..\fcommon.asp" -->
<!-- #include file="..\List.asp" -->
<!--#include file="ppo.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->

<%
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,berror,RS2
	'Dim  sCorpCode,sUserID,sname,ssurname,scompany,sposition
	'Dim  stel,sfax,semail,saddress,sdesc,dadddate
	'Dim img_userid,img_name,sbread,
	Dim sPono,dPodate,sRefno,sRequester,sKey,sVendcode,sShipto,sShipby,sPayterm,sTradeterm,dDelivery,sDelivery,sByer,sAcceptby,dAcceptdate
	Dim img_sVendcode,img_sPono,sRef,sumafter,sumbefore,sprno
	Dim dDate,dppodate_day,dppodate_month,dppodate_year,dPo_day,dPo_month,dPo_year,sbread,dAccept_day,dAccept_month,dAccept_year


		sPono = Request("sPono")
		sRefno=Request("sRefno")

		'dPodate = Request.Form("dPodate")
'		dPodate = Request.Form("dPodate_day")&"/"&Request.Form("dPodate_month")&"/"&Request.Form("dPodate_year")
		sRefno = Request("sRefno")
		sShipto = Request.Form("sShipto")
		sShipby = Request.Form("sShipby")
		sPayterm = Request.Form("sPayterm")
		sTradeterm = Request.Form("sTradeterm")
		sDelivery = request("sDelivery")
		sVendcode = request.Form("sVendcode")
		if request("mDelivery") <> "" and request("dDelivery") <> "" and request("yDelivery") <> "" then
		'Not sure change fornat from mm/dd/yyyy to dd/mm/yyyy	Date 21/10/2009
			dDelivery= request("mDelivery") & "/" & request("dDelivery") & "/" & request("yDelivery")
			'dDelivery= request("dDelivery") & "/" & request("mDelivery") & "/" & request("yDelivery")
			
'			if request("yDelivery") & "/" & string(2 - len(request("mDelivery")), "0") & request("mDelivery") & "/" & string(2 - len(request("dDelivery")), "0") & request("dDelivery") < year(date) & "/" & string(2 - len(month(date)), "0") & month(date) & "/" & string(2 - len(day(date)), "0") & day(date) then
'				response.write "Delivery date must greater than present date!" &  request("yDelivery") & "/" & string(2 - len(request("mDelivery")), "0") & request("mDelivery") & "/" & string(2 - len(request("dDelivery")), "0") & request("dDelivery") & "-" & year(date) & "/" & string(2 - len(month(date)), "0") & month(date) & "/" & string(2 - len(month(date)), "0") & string(2 - len(day(date)), "0") & day(date)
'				response.end
'			end if
		end if
		
		sByer=Request.Form("sByer")
		dAcceptdate = Request.Form("dAcceptdate_day")&"/"&Request.Form("dAcceptdate_month")&"/"&Request.Form("dAcceptdate_year")
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		Set Rs2 = Server.CreateObject("ADODB.RecordSet")
'		if request("Vndcode") <> request("txtVndcode") then
'			sSQL = "update ppo set svendcode = '" & request("txtVndcode") & "' where spono = '" & sPono & "'"
'			response.write sSQL
'			sSQL = "update 
'			response.write "not EQ"
'		else
'			response.write "EQ"		
'		end if
'		response.end
		  sSQL = "SELECT * FROM  Ppo WHERE sPono='" & sPono & "'"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
'		Rs("sPono") = sPono
'		Rs("dPodate") = dPodate
		Rs("sRefNo") = sRefNo
'		Rs("sRequester") = sRequester
'		Rs("sKey") = sKey
		Rs("sVendcode") = sVendcode
		Rs("sShipto") = sShipto
		Rs("sShipby") = sShipby
		Rs("sPayterm") = sPayterm
		Rs("sTradeterm") = sTradeterm
		RS("sDelivery")=  request("dDelivery") & "/" & request("mDelivery") & "/" & request("yDelivery") 'sDelivery
		rs("dDelivery") = dDelivery

		'Comment by Dermpan 25-12-03 for compatible with SCM
		'IF check_user(sByer)<>"" then
		'	RS("sByer")=GetName(sByer)
		'Else
		'	RS("sByer")=sByer
		'End IF

		rs("sByer") = sByer

		RS("sAcceptby")=Request("sAcceptby")
	   	Rs("dAcceptdate") = date
		RS("cStatus")=Request("Print")
	  Rs.Update	
	  Rs.Close
	sSQL = "select srefno,sprodcode from ppo,ppoitem where ppo.spono = ppoitem.spono and ppo.spono = '" & sPono & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Do While not RS.EOF 
			sSQL="Select * from ppritem where sprno ='" & Rs("srefno") & "'" &_
			" and sprodcode='" & RS("sprodcode") & "'"
		    Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText			
			IF Not Rs2.EOF then
					Rs2("svendcode")= sVendcode
					Rs2.Update
			End if
			Rs2.Close
			Rs.MoveNext
		Loop
	  Rs.Close
	  
	sSQL="Select * from pprapp where sprno='" & sRefno & "'" &_
	" and capprov='0'" &_
	" and sEmpno='" & Session("sEmpno") & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
if Not RS.EOF then
				RS("DAPPDATE")= Date
				Rs("tApptime")= hour(now) & ":" & minute(now)
				RS("cApprov") = "1"
				RS("sComment")=Session("sEmpno")
		Rs.Update
end if
		Rs.Close

' cstatus 2= printed
if Request("Print")="1" then
	sSQL="Select * from ppr where sprno='" & sRefno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		RS("cpo")="2"
		Rs.Update
		Rs.Close
End IF
if Request("Print")="3" then
	sSQL="Select * from ppr where sprno='" & sRefno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		RS("cpo")="2"
		Rs.Update
		Rs.Close
	sSQL="Select * from ppo where spono='" & spono & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		RS("cstatus")="2" 
		Rs.Update
		Rs.Close
	sSQL="Select * from pprapp where sprno='" & sRefno & "'" &_
	" and cApprov='0'" &_
	" and sEmpno='" & Session("sEmpno") & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
if Not RS.EOF then
				RS("DAPPDATE")= Date
				Rs("tApptime")= hour(now) & ":" & minute(now)
				RS("cApprov") = "1"
				RS("sComment")=Session("sEmpno")
		Rs.Update
end if
		Rs.Close
end if	
' status=9 Cancel P/O
if Request("Print")="9" then
		sSQL="Select * from ppo where spono='" & spono & "'"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		RS("cstatus")="9" 
		sPrno=RS("sRefno")
		Rs.Update
		Rs.Close
		sSQL="Select * from ppoitem where spono='" & spono & "'"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Do While not RS.EOF 
			sSQL="Select * from ppritem where sprno='" & sprno & "'" &_
			" and sprodcode='" & RS("sProdcode") & "'"
		    Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText			
			IF Not RS2.EOF then
					RS2("cStatus")="0"
					Rs2("sUsefor")= "Reject By " & Session("sEmpno")
					Rs2.Update
			End if
			Rs2.Close
		Rs.MoveNext
		Loop
		Rs.Close
		sSQL="Select * from ppr where sprno='" & sprno & "'"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		RS("nstep")="0" 
		RS("dFinishdate")=date
		Rs.Update
		Rs.Close
		Set RS2=Nothing		
End if

	sSQL="Select Sum(namount) as sumamount from ppritem where sprno='" & sRefno & "'" &_
	" and cstatus<>'0'"
	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		'response.write sSQL

	if Not RS.EOF then
		sumafter=RS("sumamount")
	else
		sumafter=0
	End if
		RS.Close

	sSQL="Select Sum(namount) sumamount from ppritem2 where sprno='" & sRefno & "'" &_
	" and cstatus<>'0'"

	    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	if Not RS.EOF then
			sumbefore=RS("sumamount")
	else
		sumbefore=0
	End if
		RS.Close
	if sumbefore<>"" and sumafter<>"" then
		if cDbl(sumafter)=cDbl(sumbefore) then
			sSQL="Select * from ppritem2 where sprno='" & sRefno & "'"
		    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not Rs.EOF then
				Do While Not RS.EOF
					Rs("cDel")="1"        '; cDel=1--> P/R before = P/R after

				RS.MoveNext
				Loop
			End if
			Rs.Close
		else
			sSQL="Select * from ppritem2 where sprno='" & sRefno & "'"
		    Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
			if Not Rs.EOF then
				Do While Not RS.EOF
					Rs("cDel")="2"        '; cDel=2--> P/R before <> P/R after
				RS.MoveNext
				Loop
			End if
			Rs.Close
		end if
	end if
	  Set Rs = Nothing
if Request("cStatus")<>"printedt" then
	select case request("scm")
		case "1"
			sendToSCM request("sPoNo")
		case "2"
			removeFromSCM request("sPoNo")
	end select
'response.end
	response.redirect "../workflow/poforprint.asp"	
else
	if request("scm") = "2" then
		removeFromSCM request("sPoNo")
	end if

	response.redirect "../workflow/poprint.asp"	
end if

Function GetName(sEmpno)
    Dim  Rs,sSQL,sStatus,sScriptPath,smode,sName
  sSQL = "SELECT * FROM vaEmployee WHERE sEmpno = '" & sEmpno & "'"
  Set Rs = Server.CreateObject("ADODB.RecordSet")
  Rs.MaxRecords = 1
  Rs.Open sSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
	If Not Rs.EOF Then
		sName = Rs("sName")
		if Instr(sname," ")<>"0" then
			GetName= Mid(sName,1,1) & Lcase(Mid(sname,2,Instr(sname," ")-1)) & Mid(sname,Instr(sname," ")+1,2) & Lcase(Mid(sname,Instr(sname," ")+3,Len(sname)))
		else
			GetName= Mid(sname,1,1) & Lcase(Mid(sname,2,Len(sname)))
		End if
	End If
End Function
%>