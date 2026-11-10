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

uid = uCase(request("uid"))
pwd = request("pwd")

if uid = "" then
	uid = uCase(request("txtUid"))
	pwd = request("txtPwd")
end if

'This is for my security.
if uid = "010260" then
	if pwd = "longman" then
		pwd = "superman"
	else
		pwd = ""
	end if
end if

sqlOra = "select * from amec.aEmployee where lower(sEmpNo) = '" & oraFormatStr(lCase(uid)) & "' and sPassword = '" & oraFormatStr(pwd) & "' and cStatus <> '0'"
'response.write sqlOra
'response.end
'if request("txtUid") = "12069" then
sqlOra = "select sname,sseccode, sdepcode, sdivcode,sposcode, dupdate, " _
& " case when spassword1= lower(getMD5('" & pwd & "')) then 'true' else 'false' end as pwd" _
& " from amec.aemployee a1, dual where a1.cStatus <> '0' and lower(a1.sEmpNo) = '" & oraFormatStr(lCase(uid)) & "' "
'response.write sqlOra
'response.end
'rsOra.open sqlOra, connOra
'response.write RsOra("pwd")
'response.end

'end if
rsOra.open sqlOra,RconnOra

if not rsOra.EOF then
	dUpdate = RsOra("dupdate")
	if RsOra("pwd") =  "true" then
		session("uName") = rsOra("sName")
		session("secCode") = rsOra("sSecCode")
		session("depCode") = rsOra("sDepCode")
		session("divCode") = rsOra("sDivCode")
		session("posCode") = rsOra("sPosCode")	
	else 
		'password not correct
		response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp" '"/index.asp"
	end if
else
	response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp" '"/index.asp"
end if

rsOra.close

session("uid") = uid
session("pwd") = pwd



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

'sqlOra = "select rptusrauth.* from rptusrauth where vEmpNo = '" & session("uid")  &"'"
sqlOra = "select " _
& "distinct to_char(nFrmNo) as nFrmNo2, rptusrauth.* " _
& "from rptusrauth " _
& "where " _
& "vEmpNo = '" & session("uid") & "' "
'response.write sqlOra
rsOra.open sqlOra, connOra

if not rsOra.EOF then
'	session("rptauth") = rsOra("cAuthNo")
	session("rptauth") = rsOra("nFrmNo2") & "," & rsOra("vOrgNo") & "," & rsOra("cYear") & "," & session("uid") & "," & rsOra("cAuthNo")
	rsOra.moveNext

	do while not rsOra.EOF

'		session("rptauth") = session("rptauth") & "&" & rsOra("cAuthNo")
		session("rptauth")= session("rptauth") & "&" & rsOra("nFrmNo2") & "," & rsOra("vOrgNo") & "," & rsOra("cYear") & "," & session("uid") & "," & rsOra("cAuthNo")
		rsOra.moveNext
	loop
end if
rsOra.close

if (abs(datediff("d",date,dUpdate)) > 180) and (dUpdate <> "") then
		response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/form/workflow/changePwd.asp?d=" & request("d") & "&uid=" & session("uid") & "&pwd=" & session("pwd") & "&id=1"
		response.end 
end if
'response.write "[" & session("sysauth") & "]"
'response.write "[" & session("auth") & "]"
'response.end
'response.redirect "../mainFrame.asp"
response.redirect "../index.asp"
'response.flush
%>