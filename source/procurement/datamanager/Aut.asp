<% 'Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: worakrit yodthonglert
	' Version: 1.00
%>	
<%
	sEmpno = Session("sEmpno")
	sSQL = "Select * From aAuthorize Where sEmpno = '"&sEmpno&"'"
	Set RS = Server.CreateObject("ADODB.RecordSet")
	Rs.Open sSQL, conn, adOpenStatic, adLockPessimistic, adCmdText
	Response.write RS("sEmpno")
	Response.write ("<BR>")
	' User Manager
	Response.write Mid(RS("sAuthor"),1,3)
	Response.write ("<BR>")
	' Ability in Datamanager
	Response.write ("Ability in Datamanager =")
	Response.write Mid(RS("sAuthor"),4,3)
	Response.write ("<BR>")
	' Ability in Vender
	Response.write (" Ability in Vender = ")
	Response.write Mid(RS("sAuthor"),7,3)
	Response.write ("<BR>")
	' Ability in Product
	Response.write (" Ability in Product = ")
	Response.write Mid(RS("sAuthor"),10,3)
	Response.write ("<BR>")
	' Ability in P/R
	Response.write (" Ability in P/R = ")
	Response.write Mid(RS("sAuthor"),13,3)
	Response.write ("<BR>")
	Response.write Mid(RS("sAuthor"),16,3)
	Response.write ("<BR>")
	Response.write Mid(RS("sAuthor"),19,3)
	Response.write ("<BR>")
	Response.write Mid(RS("sAuthor"),22,3)
	Response.write ("<BR>")
	' Ability in Report
	Response.write (" Ability in  Report = ")
	Response.write Mid(RS("sAuthor"),25,3)	
	Response.write ("<BR>")
	' Department
	Response.write (" Ability in Department = ")
	Response.write Mid(RS("sAuthor"),28,3)

	RS.Close
	Set RS = Nothing
	Response.write ("<BR>")
	Response.write Session("sAuthor")
%>