<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim uid, pwd , pcname

response.buffer = true

set rsOra = createObject("adodb.recordset")

uid = uCase(request("uid"))
pwd = request("pwd")

if uid = "" then
	uid = uCase(request("txtUid"))
	pwd = request("txtPwd")
end if



sqlOra = "select * from amec.aEmployee where lower(sEmpNo) = '" & oraFormatStr(lCase(uid)) & "' and sPassword = '" & oraFormatStr(pwd) & "' and cStatus <> '0'"
'rsOra.open sqlOra, connOra
rsOra.open sqlOra, RconnOra

if not rsOra.EOF then
	session("uName") = rsOra("sName")
	session("secCode") = rsOra("sSecCode")
	session("depCode") = rsOra("sDepCode")
	session("divCode") = rsOra("sDivCode")
	session("posCode") = rsOra("sPosCode")	
else
	response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp" '"/index.asp"
end if

rsOra.close

session("uid") = uid
session("pwd") = pwd

sqlOra = "delete From usrAuth where nfrmno = '1' and vorgno = '030101' and cyear = '17'"
rsOra.open sqlOra, connOra
sqlOra = "delete From usrAuth where nfrmno = '2' and vorgno = '030101' and cyear = '17'"
rsOra.open sqlOra, connOra

'Read forms authority
sqlOra = "select " _
& "distinct to_char(nFrmNo) as nFrmNo2, usrAuth.* " _
& "from usrAuth " _
& "where " _
& "vEmpNo = '" & session("uid") & "' "
'response.write sqlOra & "<br>"
'response.end
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	session("auth") = rsOra("nFrmNo2") & "," & rsOra("vOrgNo") & "," & rsOra("cYear") & "," & session("uid") & "," & rsOra("cAuthNo")

	rsOra.moveNext
	
	do while not rsOra.EOF
		session("auth") = session("auth") & "&" & rsOra("nFrmNo2") & "," & rsOra("vOrgNo") & "," & rsOra("cYear") & "," & session("uid") & "," & rsOra("cAuthNo")
		rsOra.moveNext
	loop
end if

rsOra.close

'Read system authority
sqlOra = "select usrSysAuth.* from usrSysAuth where vEmpNo = '" & session("uid")  &"'"
'response.write sqlOra & "<br>"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	session("sysauth") = rsOra("cAuthNo")

	rsOra.moveNext
	
	do while not rsOra.EOF
		session("sysauth") = session("sysauth") & "&" & rsOra("cAuthNo")
		rsOra.moveNext
	loop
end if

rsOra.close
'Response.Redirect "http://esrsclient/getPcname/index.php?ip=" & Request.ServerVariables("REMOTE_ADDR") & "&system=webflow"
'if Request("pcname") <> "" then 
'		pcname = Request("pcname")
'else 
'		pcname = Request.ServerVariables("REMOTE_ADDR")
'end if 

if request("txtUid") <> "" then
		Response.Redirect "http://amecweb/getPcname/index.php?ip=" & Request.ServerVariables("REMOTE_ADDR") & "&system=webflowlv"
end if

if Request("pcname") <> "" then 
		pcname = Request("pcname")
else 
		pcname = Request.ServerVariables("REMOTE_ADDR")
end if 

	
	sqlOra = "INSERT INTO amec.usrAccessLog VALUES('" & Session("uid") & "', '" & pcname & "', '" & Request.ServerVariables("REMOTE_ADDR") & "' , CURRENT_TIMESTAMP, 'WEBFLOW', '1', NULL)"
	'response.write sqlOra
	'response.end
	rsOra.open sqlOra, connOra


'response.write Request("pcname")
'response.end

response.redirect "../mainFrame.asp"
'response.flush
%>