<!--#include file="inc/system.inc"-->
<!--#include file="inc/db.inc"-->
<!--#include file="inc/connUtil.inc"-->
<%
if session("uid") = "" then
	response.redirect "loginFrm.asp"
end if
%>

<frameset rows="10%,*" frameborder="0" framespacing="0" border="0">
	<frame name="winTop" src="top.asp" scrolling="no" noresize>
	<frameset cols="115,*" framespacing="0">
		<frame name="winLeft" src="left.asp" scrolling="auto">
		<frame name="winRight" src="index.asp" scrolling="auto">
	</frameset>
</frameset>