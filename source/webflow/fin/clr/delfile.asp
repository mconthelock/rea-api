<% Option Explicit %>
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include file="../../inc/system.inc"--> 
<!--#include file="../../inc/connUtil.inc"--> 
<!--#include file="../../inc/db.inc"--> 


<%
	Dim sSQL,Rs
    Dim objFile, spath, filepath, sprno
	Dim no, orgNo, y, y2, runNo

      '''''''''''''''''''''''''''''''''Dalete File
      set objfile = CreateObject("Scripting.FileSystemObject")
      sPath = request.ServerVariables("APPL_PHYSICAL_PATH")
	  
      sPath = sPath  & FIN_SUB_SYSTEM_FOLDER & "\clr\file\" & request("frmno") & "_"  & request("orgNo") & "_" & request("y") & "_" & request("y2") & "_" &  request("runNo") & "\" & request("sfile")

      If (objfile.FileExists(sPath)) Then
		objfile.DeleteFile(sPath)
      End If

      '''''''''''''''''''''''''''''''''Delete DataBase
	  ' Create a recordset and open it for adding records
	  
	sSQL = "delete from attClrFrm where nfrmno = '" & request.querystring("frmno") & "'" _
	& " and vorgno = '" & request.querystring("orgNo") & "'" _
	& " and cyear = '" & request.querystring("y") & "'" _
	& " and cyear2 = '" & request.querystring("y2") & "'" _
	& " and nrunno = '" & request.querystring("runNo") & "'" _
	& " and sfile = '" & request.querystring("sfile") & "'"
	  'response.write sSql
	  'response.end
	  Set Rs = Server.CreateObject("ADODB.RecordSet")
	  Rs.Open sSQL, connOra
	  Set Rs = Nothing
'	  response.write Request("bp")
'	  response.end
	  Response.Redirect "viewForm.asp?no=" & request("frmno") & "&orgNo=" & Request("orgNo") & "&y=" & Request("y") & "&y2=" & Request("y2") & "&runNo=" & Request("runNo") & "&bp=" &  Request("bp")
%>


        