<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/redirectUtil.inc"-->
<%
		bp = request.querystring("bp")
		'response.write getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))
		'response.end
		if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))=MODE_ADD then
			'response.redirect ("http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER &"/"&  PS_SUB_SYSTEM_FOLDER &"/shareload/CreateForm.asp?no="& request.querystring("no")& "&orgNo="& request.querystring("orgNo")& "&y="& request.querystring("y") & "&bp=" & server.urlencode(bp))
			'response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/dedform/PCBDwg/PCBDwg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "///1")
			url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/dedform/PCBDwg/PCBDwg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "///1"

		else
			'response.redirect("http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER&"/"&  PS_SUB_SYSTEM_FOLDER &"/shareload/viewForm.asp?no="& request.querystring("no")& "&orgNo="& request.querystring("orgNo")& "&y="& request.querystring("y") & "&empNo=" & request.querystring("empNo") & "&y2=" & request.querystring("y2") & "&runNo=" & request.querystring("runNo")& "&m=" & request.querystring("m") & "&menu=" & request.querystring("menu") & "&bp=" & server.urlencode(bp))
			'response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/dedform/PCBDwg/PCBDwg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/2" )
			url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/dedform/PCBDwg/PCBDwg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/2"
		end if
        Call RedirectTo(url)
%>