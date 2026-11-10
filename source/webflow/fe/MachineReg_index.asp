<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/flowUtil.inc"-->
<!--#include file="../inc/oraUtil.inc"-->
<!--#include file="../inc/redirectUtil.inc"-->

<%	
			dim mcREG
			if InStr(SYSTEM_FOLDER,"test") > 0  then 
				mcREG = "test.mitsubishielevatorasia.co.th"
			else 
				mcREG = ".mitsubishielevatorasia.co.th"
			end if

			' --- ส่วนเช็ก HTTP / HTTPS แบบไดนามิก ---
			dim protocol
			if LCase(Request.ServerVariables("HTTPS")) = "on" or Request.ServerVariables("SERVER_PORT") = "443" or LCase(Request.ServerVariables("HTTP_X_FORWARDED_PROTO")) = "https" then
				protocol = "https://"
			else
				protocol = "http://"
			end if
    		' -----------------------------------------	
						
			dim USER_AGENT 
			USER_AGENT = Request.ServerVariables("HTTP_USER_AGENT")
			response.write(USER_AGENT)
			
			if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))=MODE_ADD then 	
				'response.redirect "http://amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/MachineReg_FORMADD/ADD/" & session("uid")   '"/index.asp"	
				
				if InStr(USER_AGENT,"Edg") > 0 or InStr(USER_AGENT,"Firefox") > 0  or InStr(USER_AGENT,"Chrome") > 0  then			   
					url = protocol & "amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/MachineReg_FORMADD/ADD/" & session("uid")
					' response.redirect ("http://amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/MachineReg_FORMADD/ADD/" & session("uid"))   '"/index.asp"	
				else
					url = "microsoft-edge:" & protocol & "amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/MachineReg_FORMADD/ADD/" & session("uid")
					' response.redirect ("microsoft-edge:http://amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/MachineReg_FORMADD/ADD/" & session("uid"))   '"/index.asp"	
				end if
			else	
				'response.redirect("http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER&"/"&  PS_SUB_SYSTEM_FOLDER &"/shareload/viewForm.asp?no="& request.querystring("no")& "&orgNo="& request.querystring("orgNo")& "&y="& request.querystring("y") & "&empNo=" & request.querystring("empNo") & "&y2=" & request.querystring("y2") & "&runNo=" & request.querystring("runNo")& "&m=" & request.querystring("m") & "&menu=" & request.querystring("menu") & "&bp=" & server.urlencode(bp))
				'response.redirect ("http://amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )			
				
				if InStr(USER_AGENT,"Edg") > 0 or InStr(USER_AGENT,"Firefox") > 0  or InStr(USER_AGENT,"Chrome") > 0  then		
					url = protocol & "amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3"
					' response.redirect ("http://amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )			
				else
					url = "microsoft-edge:" & protocol & "amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3"
					' response.redirect ("microsoft-edge:http://amecweb" & mcREG & "/webflow/feform/MachineReg/MachineReg/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )			
				end if
			end if 
		Call RedirectTo(url)
	
%>