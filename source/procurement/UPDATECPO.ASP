<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<!-- #include file ="Conn.asp" -->

<%
	' Author: Phornthep Khongsathian
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,fromday,frommonth,fromyear,today,tomonth,toyear
	  ' Create a recordset and open it for adding records
	fromday=Request("from_day")
	frommonth=Request("from_month")
	fromyear=Request("from_year")
	today=Request("to_day")
	tomonth=Request("to_month")
	toyear=Request("to_year")


	  sSQL = "DELETE FROM pInvoice"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  sSQL="Delete From pTax"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  sSQL="Delete from preceiveitem"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		sSQL="Delete from pReceive"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		sSQL="Delete from ppoitem"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		sSQL="Delete from ppo"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		sSQL="Delete from ppritem"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		sSQL="Delete from pprapp"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		sSQL="Delete from ppr"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing

	  Response.write "Delete complete"
%>
