<% Option Explicit %>
<!-- #include file="../adovbs.inc" -->
<!-- #include file ="../Conn.asp" -->

<%
	' Author: Ekkapan Kongchoo

	Dim sSQL,Rs
    Dim objFile,spath,filepath,sCorpCode
		sCorpCode = Request("sCorpCode")

      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
      sPath = request.ServerVariables("APPL_PHYSICAL_PATH")
      sPath = sPath  & "/shop/" & Request("scorpcode") & "/images/"
 
      sSQL = "SELECT sfilename FROM bimagelib WHERE sCorpCode='" & Request("sCorpCode") &"' and simgID='"&Request("simgID")&"'"   
      Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL,conn,adOpenKeySet,adLockPessimistic,adCmdText
      
      sPath = sPath & rs("sfilename")
	  Set Rs = Nothing

      'filepath = Server.MapPath(sPath)
	'	  Set objfile = CreateObject("Scripting.FileSystemObject")
      If (objfile.FileExists(sPath)) Then
		objfile.DeleteFile(sPath)
      End If

      'objfile.DeleteFile(sPath) 


      '''''''''''''''''''''''''''''''''Delete DataBase
	  ' Create a recordset and open it for adding records
	  sSQL = "DELETE FROM bimagelib WHERE sCorpCode='" & Request("sCorpCode") &"' and simgID='"&Request("simgID")&"'"
	  'response.write sSql
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, conn, adOpenKeySet, adLockPessimistic, adCmdText
	  Set Rs = Nothing
	  Response.Redirect "bbimagelist.asp?sCorpCode="&sCorpCode
%>


        