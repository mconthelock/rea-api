<% Option Explicit %>

<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
Dim RS,sSQL,sPrno,Counter,Maxnitem,i,berror,sErrImage,srecno,drecdate,srecby,sreceiver,sUpdateMode
Dim img_srecno,RS2,RS3

	bError = False
  	sErrImage = "<IMG SRC='.../images/icerror.gif'>"
'	response.write sSQL


		srecno = Request("srecno")
		sprno = Request("sprno")
		'######		06/08/2009	change format date from mm/dd/yyyy to dd/mm/yyyy	#####
'''		drecdate = Request.Form("drecdate_month")&"/"&Request.Form("drecdate_day")&"/"&Request.Form("drecdate_year")
		drecdate = Request.Form("drecdate_day")&"/"&Request.Form("drecdate_month")&"/"&Request.Form("drecdate_year")
		srecby= Request.Form("srecby")
		sreceiver= Request.Form("sreceiver")

	  Set Rs2 = Server.CreateObject("ADODB.RecordSet")
		sSQL="Select * from pReceive where sRecno='" & sRecno & "'"
	  Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	' Check if the End date is a valid date provided it is not empty
IF RS2.EOF then
	If (Trim(srecno) = "") Then
		bError = True
		img_srecno = sErrImage
	End If


	If bError = False Then
	  ' Create a recordset
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
		
		' Depending on the Mode, open the Recordset in eith Add or Edit mode

		If sUpdateMode = "EDIT" Then
		  sSQL = "SELECT * FROM  Preceive WHERE srecno='" & srecno &"'" 
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		Else
		  Rs.Open "Preceive", conn, adOpenKeySet, adLockPessimistic, adCmdTable
		  Rs.AddNew
		End If

		Rs("srecno") = srecno
		Rs("sprno") = sprno
		Rs("drecdate") = drecdate
		Rs("srecby") = srecby
		Rs("sreceiver") = sreceiver
	  Rs.Update	

	  Rs.Close
	  Set Rs = Nothing

	else
		'Show_Form()
	End If	' if Error 
End if

	Set Rs=Server.CreateObject("ADODB.Recordset")
	Set Rs2=Server.CreateObject("ADODB.Recordset")
	sSQL = "SELECT  * FROM ppritem  where sPrno='"&sPrno&"'" &_
	"  Order By nItem"
  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

IF Not RS.EOF then
	Counter=0
	Do While Not Rs.EOF
		Counter=Counter+1
		IF RS("cStatus") <> "0" and RS("cStatus") <> "2" and RS("cStatus") <> "3"  then
			RS("cStatus")=Request(Counter)
			RS.Update
		End if
	Rs.MoveNext
	Loop
 RS.Close
	sSQL = "SELECT  * FROM ppritem  where sPrno='"&sPrno&"'" &_
		" and cStatus= '1'" 
  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

IF Not RS.EOF then
'cpo=3 -> Received
Else
			sSQL="Select * from ppr where sprno='" & sPrno & "'"
 			 Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		if Not RS2.EOF then
			Do while Not RS2.EOF	
				RS2("cPo")="3"
				RS2.Update
			RS2.MoveNext
			Loop
		End if
	Rs2.Close
End if
	sSQL="Select * from ppo where sRefno='" & sPrno & "'"
	 Rs2.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText	
	Do while Not RS2.EOF
		RS2("creceive")="1"
		RS2("cInvoice")="1"
		RS2("cTax")="1"
		RS2.Update
		Rs2.Movenext
	Loop
		RS2.Close
	Response.redirect "../workflow/workflow.asp"
else
End If

%>