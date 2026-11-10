<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->
<!--#include file="../inc/authen.inc"-->
<%
Dim rsOra, sqlOra
Dim uid, pwd
Dim pwd2, pwd3 , newPwd
Dim connSQLSvr
Dim rsSQLSvr, sqlSQLSvr
set rsSQLSvr = createObject("adodb.recordset")	 
set connSQLSvr = createObject("adodb.connection")
connSQLSvr.open SQLSVR_CONN_STR

response.buffer = true
set rsOra = createObject("adodb.recordset")
if request("d") <> "" then
	d= request("d")
else
	d= "http://" & request.serverVariables("SERVER_NAME") & "/form/index.asp"
end if
'response.write "[" & d & "]"
'response.end

if request("txtNewPwd") <> "" then
	'sqlOra = "select spassword from amec.aemployee where lower(sEmpNo) = '" & oraFormatStr(lCase(session("uid"))) & "'" 
	sqlOra = "select spassword, spassword1, spassword2 from amec.aemployee where lower(sEmpNo) = '" & oraFormatStr(lCase(session("uid"))) & "' and nvl(spassword, 0)  <> '" & request("txtNewPwd") & "' and nvl(spassword2, 0)  <> lower(getMD5('" & request("txtNewPwd") & "')) and nvl(spassword3, 0)  <> lower(getMD5('" & request("txtNewPwd") & "'))"
	'response.write sqlOra
	rsOra.open sqlOra, connOra
	if request("txtNewPwd") = session("uid") then
		response.write "Your password cannot repeat your employee code"
		response.end
	end if
	
	if request("txtNewPwd") <> request("txtreNewPwd") then
		response.write "please type the same new password into the Confirm "
		response.end
	end if
	
	if not rsOra.EoF then 
		pwd2 = rsOra("spassword1")
		pwd3 = rsOra("spassword2")
		if request("txtOldPwd") <> rsOra("spassword") then
			response.write "old password not correct"
			response.end
		end if
	else
		response.write "Your password cannot repeat your previous passwords"
		response.end
	end if 
	
	
	rsOra.close
	
	sqlOra = "update amec.aEmployee set spassword='" & oraFormatStr(request("txtNewPwd")) & "'" _
	& " , dupdate = to_date('" & day(date) & "/" & month(date) & "/" & year(date) & "','dd/mm/yyyy') " _ 
	& " where lower(sEmpNo) = '" & oraFormatStr(lCase(session("uid"))) & "'" 
	
	sqlOra = "update amec.aEmployee set spassword='" & oraFormatStr(request("txtNewPwd")) & "', spassword1=lower(getMD5('" & oraFormatStr(request("txtNewPwd")) & "'))" _
	& " , spassword2 = '" & pwd2 & "', spassword3 = '" & pwd3 & "'" _
 	& " , dupdate = to_date('" & day(date) & "/" & month(date) & "/" & year(date) & "','dd/mm/yyyy') " _ 
	& " where lower(sEmpNo) = '" & oraFormatStr(lCase(session("uid"))) & "'" 
	'response.write sqlOra
	'response.end
	rsOra.open sqlOra, connOra
	'response.write "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/workflow/WaitApv.asp"
	'response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/workflow/changePwd.asp"
	'	response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/default.asp" '"/index.asp"
	
	sqlOra = "select spassword1 from amec.aemployee where lower(sEmpNo) = '" & oraFormatStr(lCase(session("uid"))) & "'"
	rsOra.open sqlOra, connOra
	if not rsOra.EoF then 
		newPwd = rsOra("spassword1")
	end if
	rsOra.close
	
	sqlSQLSvr = "update tb_sqlamecuser set passEnc = '" & newPwd & "' where lower(sempno) = '" & oraFormatStr(lCase(session("uid"))) & "'"
	rsSQLSvr.open sqlSQLSvr, connSQLSvr
	
	'response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/mainFrame.asp"
	'response.redirect "../mainFrame.asp"
	'response.write "[" & request("d") & "]"
	'response.end
	session("pwd") = request("txtreNewPwd")
	'response.write request("d") & "-" & session("uid") 
	'response.end

end if
%>
<script language="javascript">
	alert("password change successful");
//	window.location.href = 'http://<%=request.serverVariables("SERVER_NAME")%>/form/workflow/changePwd.asp?d=<%=request("d")%>&uid=<%=session("uid")%>&pwd=<%=session("pwd")%>'

	window.location.href = '<%=d%>?uid=<%=session("uid")%>&pwd=<%=session("pwd")%>&txtOldPwd=<%=request("txtNewPwd")%>'
</script>
<%
response.end

'if session("uid") = request("txtreNewPwd") then
'if session("pwd") = request("txtreNewPwd") then
	'response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/form/workflow/changePwd.asp?d=" & request("d") & "&uid=" & session("uid") & "&pwd=" & session("pwd")
	 '"/index.asp"
'	response.end
'else
'	response.redirect request("d") & "?uid=" & session("uid") & "&pwd=" & session("pwd")
'end if

%>

