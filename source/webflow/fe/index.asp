<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/flowUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->
<!--#include file="../inc/redirectUtil.inc"-->
<%	
		menu = request.querystring("menu")
		'response.redirect("http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/workfactory/index/" & session("uid") & "/" & menu )
		url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/workfactory/index/"  & session("uid") & "/" & menu
		Call RedirectTo(url)
%>