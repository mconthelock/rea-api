<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/flowUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->
#include file="../inc/authen.inc

<%	
			dim mcREG
			if InStr(SYSTEM_FOLDER,"test") > 0  then 
				mcREG = "test2.mitsubishielevatorasia.co.th"
			else 
				mcREG = ".mitsubishielevatorasia.co.th"
			end if
						
			
			if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))=MODE_ADD then 	
				response.redirect "http://amecweb" & mcREG & "/webflow/feform/MachineReq/MachineReq/MachineReq_FORMADD/ADD/" & session("uid")   '"/index.asp"	
			else	
				'response.redirect("http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER&"/"&  PS_SUB_SYSTEM_FOLDER &"/shareload/viewForm.asp?no="& request.querystring("no")& "&orgNo="& request.querystring("orgNo")& "&y="& request.querystring("y") & "&empNo=" & request.querystring("empNo") & "&y2=" & request.querystring("y2") & "&runNo=" & request.querystring("runNo")& "&m=" & request.querystring("m") & "&menu=" & request.querystring("menu") & "&bp=" & server.urlencode(bp))
				response.redirect ("http://amecweb" & mcREG & "/webflow/feform/MachineReq/MachineReq/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )			
			end if 
	
%>