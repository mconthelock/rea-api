<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/flowUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->
<!--#include file="../../inc/redirectUtil.inc"-->
<%	
		menu = request.querystring("menu")
		url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/WasteDisaposal/show_form_master/"  & session("uid") & "/" & menu
		Call RedirectTo(url)
%>