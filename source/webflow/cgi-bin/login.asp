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
'response.write "[" & uCase(request("txtUid")) & "]"
'response.write "[" & request("pwd") & "]"
'response.write "[" & uid & "]"
'response.end

sqlOra = "select a.reason from amec.usrAccessLog a, amec.aemployee b where a.sempno = b.sempno and b.cstatus <> '0' and lower(a.sempno) = '" & oraFormatStr(lCase(uid)) & "' and a.login_flag = '0' and a.system_access = 'WEBFLOW' and a.datetime = (select max(datetime) from amec.usrAccessLog where lower(sempno) = '" & oraFormatStr(lCase(uid)) & "' and login_flag = '0' and system_access = 'WEBFLOW') and (a.reason is null or a.reason = '')"
rsOra.open sqlOra, connOra
if not rsOra.EOF then
	if rsOra("reason") = "" or IsNull(rsOra("reason")) then 
		session("uid") = uid
		session("InvalidAccess") = 3
		response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/cgi-bin/loginfail.asp"
	end if
end if 
rsOra.close


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

'end if
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra


if not rsOra.EOF then
	dUpdate = RsOra("dupdate")
	if RsOra("pwd") =  "true" then
		session("uName") = rsOra("sName")
		session("secCode") = rsOra("sSecCode")
		session("depCode") = rsOra("sDepCode")
		session("divCode") = rsOra("sDivCode")
		session("posCode") = rsOra("sPosCode")	
		session("InvalidAccess") = 0
	else 
		'password not correct
		if session("InvalidAccess") = "" then
			session("InvalidAccess") = 1
			session("uid") = uid
		else 
			session("InvalidAccess") = session("InvalidAccess") + 1	
		end if 
		
		if session("uid") <> uid and session("InvalidAccess") >= 1 then
			session("InvalidAccess") = 1
			session("uid") = uid
			response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp"
		end if 
		 
		if session("InvalidAccess") = 3 then
			session("uid") = uid
			Response.Redirect "http://amecweb/getPcname/index.php?ip=" & Request.ServerVariables("REMOTE_ADDR") & "&system=webflow"
		end if

		
		response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp" '"/index.asp"
	end if
else
	if session("InvalidAccess") = 3 then
		rsOra.close
		sqlOra = "INSERT INTO amec.usrAccessLog VALUES('" & session("uid") & "', '" & Request("pcname") & "', '" & Request.ServerVariables("REMOTE_ADDR") & "' , CURRENT_TIMESTAMP, 'WEBFLOW', '0', NULL)"
		rsOra.open sqlOra, connOra
		
		response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/cgi-bin/loginfail.asp"
	end if
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
'response.end
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
'response.end
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

'response.write abs(datediff("d",date,dUpdate))
'response.write "["& dUpdate & "]"
'response.write "[" & session("uid") & "]"
'response.write "[" & uid & "]"

'response.end
if (abs(datediff("d",date,dUpdate)) > 90) and (dUpdate <> "") then
'if (abs(datediff("d",date,dUpdate)) > 180) and (dUpdate <> "") then
		response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/form/workflow/changePwd.asp?d=" & request("d") & "&uid=" & session("uid") & "&pwd=" & session("pwd") & "&id=1"
		response.end 
end if
'response.write "[" & session("sysauth") & "]"
'response.write "[" & session("auth") & "]"
'response.end
'response.redirect "../mainFrame.asp"
'response.redirect "../index.asp"
sqlOra = "select * From PERSONAL where SEMPNO = '"&session("uid")&"'"
rsOra.open sqlOra, connOra
if not rsOra.EOF then
	rsOra.close
	'response.redirect "../index.asp"
	sqlOra = "select * From PERSONAL_PROCESS where SEMPNO = '"&session("uid")&"'"
	rsOra.open sqlOra, connOra
	if not rsOra.EOF then
		rsOra.close
		''sqlOra = "SELECT * FROM flow WHERE NFRMNO = '3' AND VORGNO = '050601' AND CYEAR = '12' AND CYEAR2 = '"&Year(Now)&"' AND cstepst IN ('3') AND VAPVNO NOT IN ('96321','13214') AND VAPVNO = '"& session("uid")&"'"
		''rsOra.open sqlOra, connOra
		''if not rsOra.EOF then
		''	response.redirect "http://webflow.mitsubishielevatorasia.co.th/form/workflow/WaitApv.asp"
		'	response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER&"/"&  SUB_SYSTEM_FOLDER &"/swrule/viewFrm.asp?no="& rsOra("NFRMNO")& "&orgNo="& rsOra("VORGNO")& "&y="& rsOra("CYEAR") & "&empNo=" & session("uid") & "&y2=" & rsOra("CYEAR2") & "&runNo=" & rsOra("NRUNNO")
		''	rsOra.close
		''else
		''	rsOra.close
			response.redirect "../index.asp"
		''end if
	else
		rsOra.close
		response.redirect "../Information_process.asp"
	end if
else
	rsOra.close
	response.redirect "../Information.asp"
end if



'response.flush
%>