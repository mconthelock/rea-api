<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../fcommon.asp" -->
<!-- #include file="../security/secure.asp" -->

<%
	' Author: Worakrit   yodthonglert
	' Version: 1.01
%>	

<%
   Dim Rs,sSQL,RS2,sSQL2
	Dim sExptype
	Dim sEmpno, sDivcode, sDepcode, sSeccode
	Dim SectionCode
	Dim sFlow, NStep, sStep
	Dim MaxStep, Counter
	Dim App01, App02, App03
	Dim sMode
	Dim sFlow

	sEmpno = Session("sEmpno")
	sPoscode = Session("sPoscode")
	sDivcode = Session("sDivcode")
	sDepcode = Session("sDepcode")
	sSeccode = Session("sSeccode")
	sName = Session("sName")
	Counter=1
	
	sExptype = Request("sExptype")
%>

<%
	' Check Seccode in pAppflow.sLevel find sFlow to Create pPrapp
	sSQL = "SELECT * From sAppFlow"
	Set Rs = Server.CreateObject("ADODB.RecordSet")
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	While Not RS.EOF

	SectionCode = Mid(RS("sLevel"),12,5)
	IF SectionCode = sSeccode Then
		sFlow = RS("sFlow")
	End IF

	RS.MoveNext
	Wend
	RS.Close
	
%> 

<%

	sSQL = "SELECT * From sAppStep where sFlow ="&sFlow
	Set Rs = Server.CreateObject("ADODB.RecordSet")
 	Rs.Open sSQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
	While Not Counter = 4
	Counter = Counter + 1
		
		


	RS.MoveNext
	Wend
	RS.Close

%>


