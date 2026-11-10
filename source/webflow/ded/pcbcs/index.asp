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
			'response.redirect ("http://amecweb/webflow/dedform/PCB_CheckSheet/pcb_checksheet/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "//")
			url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/dedform/PCB_CheckSheet/pcb_checksheet/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "//"
		else
			'response.redirect ("http://amecweb/webflow/dedform/PCB_CheckSheet/pcb_checksheet/showform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo"))
			url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/dedform/PCB_CheckSheet/pcb_checksheet/showform/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo")
		end if
		Call RedirectTo(url)
%>