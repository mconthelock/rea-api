<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/flowUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->

<%	
		menu = request.querystring("menu")
		response.redirect("http://amecwebtest/webflow/feform/mfg/WasteDisaposal/show_form_master/" & session("uid") & "/" & menu )
		
%>