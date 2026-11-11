<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<!-- #include file ="Conn.asp" -->
<!-- #include file ="fcommon.asp" -->

<%
	' Author: Phornthep Khongsathian
	' Version: 1.01
%>	

<%
if Request("sDelAll")=1 then
	Deleteall()
else
	Delete()
End if
Sub Delete()
	Dim sSQL,Rs,fromday,frommonth,fromyear,to_day,tomonth,toyear,fromdate,todate
	  ' Create a recordset and open it for adding records
	fromday=Request("from_day")
	frommonth=Request("from_month")
	fromyear=Request("from_year")
	to_day=Request("to_day")
	tomonth=Request("to_month")
	toyear=Request("to_year")
	fromdate=frommonth & "/" & fromday & "/" & fromyear
	todate=tomonth & "/" & to_day & "/" & toyear
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	sSQL="Select * FROM pInvoice where srecno in" &_
		" (Select srecno from preceive where sprno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate<= " & to_date(todate) & "))"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	If Not RS.EOF then
		sSQL="Select * FROM pTax where srecno in" &_
			" (Select srecno from preceive where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate<= " & to_date(todate) & "))"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		IF Not RS.EOF then
		  sSQL = "DELETE FROM pInvoice where srecno in" &_
			" (Select srecno from preceive where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate<= " & to_date(todate) & "))"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		  sSQL="Delete From pTax where srecno in" &_
			" (Select srecno from preceive where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & "))"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		  sSQL="Delete from preceiveitem where srecno in" &_
			" (Select srecno from preceive where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & "))"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

			sSQL="Delete from pReceive where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & ")"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

			sSQL="Delete from ppoitem where spono in" &_
			" (Select spono from ppo where srefno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & "))"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

			sSQL="Delete from ppo where srefno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & ")" 
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppritem where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & ")"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppritem2 where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & ")"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

			sSQL="Delete from pprapp where sprno in" &_
			" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate) & ")"
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

			sSQL="Delete from ppr where  dprdate >= " & to_date(fromdate) &_
			" and dprdate <= " & to_date(todate)
		  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		  Set Rs = Nothing

		  Response.write "Delete complete"
	else
	  Response.write "Delete Don't complete"
	End if
else
	  Response.write "Delete Don't complete"
End if
End Sub
%>
<%
Sub DeleteAll()
	Dim sSQL,Rs,fromday,frommonth,fromyear,to_day,tomonth,toyear,fromdate,todate
	  ' Create a recordset and open it for adding records
	fromday=Request("from_day")
	frommonth=Request("from_month")
	fromyear=Request("from_year")
	to_day=Request("to_day")
	tomonth=Request("to_month")
	toyear=Request("to_year")
	fromdate=frommonth & "/" & fromday & "/" & fromyear
	todate=tomonth & "/" & to_day & "/" & toyear
	  sSQL = "DELETE FROM pInvoice where srecno in" &_
		" (Select srecno from preceive where sprno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate<= " & to_date(todate) & "))"

	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL="Delete From pTax where srecno in" &_
		" (Select srecno from preceive where sprno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate) & "))"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL="Delete from preceiveitem where srecno in" &_
		" (Select srecno from preceive where sprno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate) & "))"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from pReceive where sprno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate) & ")"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppoitem where spono in" &_
		" (Select spono from ppo where srefno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate) & "))"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppo where srefno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate) & ")" 
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppritem where sprno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate) & ")"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from pprapp where sprno in" &_
		" (Select sprno from ppr where dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate) & ")"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppr where  dprdate >= " & to_date(fromdate) &_
		" and dprdate <= " & to_date(todate)
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing

	  Response.write "Delete complete"
End Sub
%>

