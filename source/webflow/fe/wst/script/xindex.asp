<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<%	
		bp = request.querystring("bp")
		if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))=MODE_ADD then 	
			response.redirect ("http://amecwebtest/webflow/feform/mfg/WasteDisposal/createform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "//")			
		else
			response.redirect ("http://amecwebtest/webflow/feform/mfg/WasteDisposal/showform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") & "/" & getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid")) )	
		end if 
%>

