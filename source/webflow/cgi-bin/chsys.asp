<%
'response.write session("uid") & "<br>"
'response.write session("pwd")
'response.write request("d")
'response.end
if session("uid") = session("pwd") then
	response.redirect "http://" & request.serverVariables("SERVER_NAME") & "/form/workflow/changePwd.asp?d=" & request("d") & "&uid=" & session("uid") & "&pwd=" & server.urlEncode(session("pwd"))
	'"/index.asp"
	response.end
else
	response.redirect request("d") & "?uid=" & session("uid") & "&pwd=" & server.urlEncode(session("pwd"))
end if

%>