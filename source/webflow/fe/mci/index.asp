<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/redirectUtil.inc"-->
<%	
		bp = request.querystring("bp")		
		if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))= MODE_ADD then 	
			url ="http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/workfactory/req_inqpec/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") 
			' response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/workfactory/req_inqpec/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") )		
		else	
		url ="http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/workfactory/show_form_approve_ins/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") 
			' response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/feform/mfg/workfactory/show_form_approve_ins/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") )			
		end if 
		
		Call RedirectTo(url)
%>