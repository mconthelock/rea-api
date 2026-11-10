<%@EnableSessionState=True %>
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->
<%

Dim rsOra, sqlOra
Dim uid, pwd
Dim dUpdate
response.buffer = true
set rsOra = createObject("adodb.recordset")
if(request("ptype") = "N" ) then
	Session("sEmpno") = ""
	session.abandon
else
	sqlOra = "insert into PERSONAL_PROCESS(SEMPNO,ACCEPTDATE,IP) values('"&session("uid")&"' , CURRENT_TIMESTAMP ,'" & Request.ServerVariables("REMOTE_ADDR") & "')"
	rsOra.open sqlOra, connOra
end if
response.redirect "/" & SYSTEM_FOLDER & "/index.asp"
%>