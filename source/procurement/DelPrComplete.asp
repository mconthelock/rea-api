<% Option Explicit %>
<!-- #include file="adovbs.inc" -->
<!-- #include file ="Conn.asp" -->
<!-- #include file ="fcommon.asp" -->

<%
	' Author: Phornthep Khongsathian
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs,fromday,frommonth,fromyear,to_day,tomonth,toyear,fromdate,todate,sprno
	  ' Create a recordset and open it for adding records
	sprno=Request("sPrno")
	fromday=Request("from_day")
	frommonth=Request("from_month")
	fromyear=Request("from_year")
	to_day=Request("to_day")
	tomonth=Request("to_month")
	toyear=Request("to_year")
	fromdate=frommonth & "/" & fromday & "/" & fromyear
	todate=tomonth & "/" & to_day & "/" & toyear
	  sSQL = "DELETE FROM pInvoice where srecno in" &_
		" (Select srecno from preceive where sprno ='" & sprno & "')"

	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL="Delete From pTax where srecno in" &_
		" (Select srecno from preceive where sprno ='" & sprno & "')"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

	  sSQL="Delete from preceiveitem where srecno in" &_
		" (Select srecno from preceive where sprno ='" & sprno & "')"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from pReceive where sprno ='" & sprno & "'"

	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppoitem where spono in" &_
		" (Select spono from ppo where srefno ='" & sprno & "')"

	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppo where srefno ='" & sprno & "'"

	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppritem where sprno ='" & sprno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from pprapp where sprno ='" & sprno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText

		sSQL="Delete from ppr where sprno='" & sprno & "'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing

	  Response.write "Delete complete"
%>
<%
Sub Delfile()
	Dim sSQL,Rs
    Dim objFile,spath,filepath,sprno,nitem,sitemno
		sprno = Request("sprno")

      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
 
      sSQL = "SELECT sfile FROM ppritemdesc WHERE sprno ='" & sprno & "'"
 
      Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
	Do While Not RS.EOF      
	      sPath = request.ServerVariables("APPL_PHYSICAL_PATH")
	      sPath = sPath  & "procurement/workflow/ppritemdescimages/"
    	  sPath = sPath & rs("sfile")

	      'filepath = Server.MapPath(sPath)
		'	  Set objfile = CreateObject("Scripting.FileSystemObject")
	      If (objfile.FileExists(sPath)) Then
			objfile.DeleteFile(sPath)
	      End If
	RS.MoveNext
	Loop

	  Set Rs = Nothing

      'objfile.DeleteFile(sPath) 


      '''''''''''''''''''''''''''''''''Delete DataBase
	  ' Create a recordset and open it for adding records
	  sSQL = "DELETE FROM ppritemdesc WHERE  sprno ='" & sprno & "'"

	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
'	  Response.Redirect "Showprdesc.asp?sprno=" & sprno & "&nitem=" & Request("nitem") & "&sExptype=" & Request("sExptype")
End Sub
%>

