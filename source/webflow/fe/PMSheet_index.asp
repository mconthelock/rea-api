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
						
			dim USER_AGENT 
			USER_AGENT = Request.ServerVariables("HTTP_USER_AGENT")
			response.write(USER_AGENT)
			
			' if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))=MODE_ADD then 	
			' 	if InStr(USER_AGENT,"Edg") > 0 or InStr(USER_AGENT,"Firefox") > 0  or InStr(USER_AGENT,"Chrome") > 0  then			   
			' 		url ="http://amecweb" & mcREG & "/webflow/feform/PMSheet/PMSheet/MachineReg_FORMADD/ADD/" & session("uid") 
			' 		' response.redirect ("http://amecweb" & mcREG & "/webflow/feform/PMSheet/PMSheet/MachineReg_FORMADD/ADD/" & session("uid") )  '"/index.asp"	
			' 	else
			' 		url ="microsoft-edge:http://amecweb" & mcREG & "/webflow/feform/PMSheet/PMSheet/MachineReg_FORMADD/ADD/" & session("uid") 
			' 		' response.redirect ("microsoft-edge:http://amecweb" & mcREG & "/webflow/feform/PMSheet/PMSheet/MachineReg_FORMADD/ADD/" & session("uid") )  '"/index.asp"	
			' 	end if
			' else	
				if InStr(USER_AGENT,"Edg") > 0 or InStr(USER_AGENT,"Firefox") > 0  or InStr(USER_AGENT,"Chrome") > 0  then		   
					url ="http://amecweb" & mcREG & "/smmt-2022/Machine/PMSheetGetViewWebflow/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" 
					' response.redirect ("http://amecweb" & mcREG & "/smmt-2022/Machine/PMSheetGetViewWebflow/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )
					' 'response.redirect ("http://amecweb" & mcREG & "/webflow/feform/PMSheet/PMSheet/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )			
				else
					url ="microsoft-edge:http://amecweb" & mcREG & "/smmt-2022/Machine/PMSheetGetViewWebflow/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" 
					' response.redirect ("microsoft-edge:http://amecweb" & mcREG & "/smmt-2022/Machine/PMSheetGetViewWebflow/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )	
					' 'response.redirect ("microsoft-edge:http://amecweb" & mcREG & "/webflow/feform/PMSheet/PMSheet/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/3" )			
				end if
			' end if 
		Call RedirectTo(url)
	
%>