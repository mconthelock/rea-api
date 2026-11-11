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
	  Set Rs = Server.CreateObject("ADODB.RecordSet")

	  delfile()
	  
	  sSQL = "DELETE FROM pprpackage WHERE sprno='" & Request.QueryString("sprno") &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  
	  sSQL = "DELETE FROM ppritem WHERE sprno='" & Request.QueryString("sprno") &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
		
	  sSQL = "DELETE FROM ppr WHERE sprno='" & Request.QueryString("sprno") &"'"
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "../workflow/workflow.asp"
%>
<%
Sub Delfile()
	Dim sSQL,Rs
    Dim objFile,spath,filepath,sprno,nitem,sitemno
		sprno = Request("sprno")

      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
 
      sSQL = "SELECT sfile FROM ppritemdesc WHERE  sprno='" & Request.QueryString("sprno") &"'"
 
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
	  sSQL = "DELETE FROM ppritemdesc WHERE  sprno='" & Request.QueryString("sprno") &"'"

	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
'	  Response.Redirect "Showprdesc.asp?sprno=" & sprno & "&nitem=" & Request("nitem") & "&sExptype=" & Request("sExptype")
End Sub
%>