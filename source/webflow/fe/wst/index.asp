<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/redirectUtil.inc"-->
<%	
		bp = request.querystring("bp")
		if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))=MODE_ADD then 	
			url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/WasteDisaposal/showcreateform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "//"
			' response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/WasteDisaposal/showcreateform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "//")			
		else
			 url ="http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/WasteDisaposal/showform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") & "/" & getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid")) 
			' response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/WasteDisaposal/showform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") & "/" & getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid")) )	
		end if 
		Call RedirectTo(url)
%>

