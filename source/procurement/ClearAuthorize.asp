<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<!-- #include file ="Conn.asp" -->

<%
	' Author: Phornthep Khongsathian
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs
	  ' Create a recordset and open it for adding records
	  sSQL = "Update Aauthorize set sAuthor='000000000000000000000000000000000000' where sEmpno<>'SYSTEM'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Response.write "Clear complete"
%>
