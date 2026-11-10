<!--#include file="inc/login.inc"-->


<%
response.buffer = true
response.write "<font color=#0000FF>Start " & time
for i = 1 to 1000000
	if i mod 50000 = 0 then
		response.write "."
		response.flush
	end if
next
response.write " End " & time & "</font>"
%>