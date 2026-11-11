<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->
<!-- #include file="../security/secure.asp" -->
<%
	' Author: Ekkapan Kongchoo
	' Version: 1.01
%>	

<%
	Dim sSQL,Rs
	  ' Create a recordset and open it for adding records
	Delfile()
	  sSQL = "DELETE FROM ppritem WHERE sprno='" & Request.QueryString("sprno") &"'" &_
		"and nitem='" & Request.QueryString("nitem") & "'"
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "../workflow/showpr.asp?sPrno=" & Request.QueryString("sprno")
%>
<%
Sub Delfile()
	Dim sSQL,Rs
    Dim objFile,spath,filepath,sprno,nitem,sitemno
		sprno = Request("sprno")

      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
      sPath = request.ServerVariables("APPL_PHYSICAL_PATH")
      sPath = sPath  & "procurement/workflow/ppritemdescimages/"
 
      sSQL = "SELECT sfile FROM ppritemdesc WHERE sprno='" & Request("sprno") &"'" &_
		" and nitem='" & Request("nitem") & "'"
      Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
 	if Not Rs.EOF then
      sPath = sPath & rs("sfile")


      'filepath = Server.MapPath(sPath)
	'	  Set objfile = CreateObject("Scripting.FileSystemObject")
      If (objfile.FileExists(sPath)) Then
		objfile.DeleteFile(sPath)
      End If
	End if
	  Rs.Close
	  Set Rs = Nothing

      'objfile.DeleteFile(sPath) 


      '''''''''''''''''''''''''''''''''Delete DataBase
	  ' Create a recordset and open it for adding records
	  sSQL = "DELETE FROM ppritemdesc WHERE  sprno='" & Request("sprno") &"'" &_
		" and nitem='"&Request("nitem")&"'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	'  Response.Redirect "ppritemedt2.asp?sprno=" & sprno & "&nitem=" & Request("nitem") & "&sExptype=" & Request("sExptype")
End Sub
%>