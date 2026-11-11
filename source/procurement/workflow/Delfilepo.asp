<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->

<%
	Dim sSQL,Rs
    Dim objFile,spath,filepath,sprno,nitem,sitemno,spono
		sprno = Request("sprno")
		spono = Request("spono")
      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
      sPath = request.ServerVariables("APPL_PHYSICAL_PATH")
      sPath = sPath  & "procurement/workflow/ppritemdescimages/"
 
      sSQL = "SELECT sfile FROM ppritemdesc WHERE sprno='" & Request("sprno") &"'" &_
		" and nitem='" & Request("nitem") & "'" &_
		" and sitemno='"&Request("sitemno") &"'"   
      Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
      
      sPath = sPath & rs("sfile")
	  Set Rs = Nothing

      'filepath = Server.MapPath(sPath)
	'	  Set objfile = CreateObject("Scripting.FileSystemObject")
      If (objfile.FileExists(sPath)) Then
		objfile.DeleteFile(sPath)
      End If

      'objfile.DeleteFile(sPath) 


      '''''''''''''''''''''''''''''''''Delete DataBase
	  ' Create a recordset and open it for adding records
	  sSQL = "DELETE FROM ppritemdesc WHERE  sprno='" & Request("sprno") &"'" &_
		" and nitem='"&Request("nitem")&"'" &_
		" and sitemno='"&Request("sitemno")&"'"   
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "ppo.asp?sPono=" & spono
%>


        