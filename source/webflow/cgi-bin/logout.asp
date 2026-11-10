<!--#include virtual="/form/inc/system.inc"-->
<meta http-equiv="Pragma" content="no-cache" >
<meta http-equiv="Expires" content="0">
<%
session.abandon
response.redirect "/" & SYSTEM_FOLDER & "/default.asp" '"/index.asp"
%>