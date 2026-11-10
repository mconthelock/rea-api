<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/redirectUtil.inc"-->
<%
		bp = request.querystring("bp")
		'response.write "In the process new system"
		'response.end
		  dim sqlcj
		  dim rschkj
		  set rschkj=Createobject("Adodb.recordset")

		if getMode(request.querystring("no"), request.querystring("orgNo"),request.querystring("y"),request.querystring("y2"), request.querystring("runNo"),session("uid"))=MODE_ADD then
			'response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/finform/clr/Clr/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "///1")
			'response.redirect ("http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER &"/"&  FIN_SUB_SYSTEM_FOLDER &"/clr/CreateForm.asp?no="& request.querystring("no")& "&orgNo="& request.querystring("orgNo")& "&y="& request.querystring("y") & "&bp=" & server.urlencode(bp))
            url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/finform/clr/Clr/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "///1"
		else
		    sqlcj = "select dreqdate from form where NFRMNO = '"&request.querystring("no")&"' and VORGNO = '"&request.querystring("orgNo")&"' and CYEAR = '"&request.querystring("y")&"' and CYEAR2 = '"&request.querystring("y2")&"' and NRUNNO = '"&request.querystring("runNo")&"'"
			rschkj.Open sqlcj, connOra
			if datevalue(rschkj("dreqdate")) < datevalue("02/05/2025") then
				'response.redirect("http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER&"/"&  FIN_SUB_SYSTEM_FOLDER &"/clr/viewForm.asp?no="& request.querystring("no")& "&orgNo="& request.querystring("orgNo")& "&y="& request.querystring("y") & "&empNo=" & request.querystring("empNo") & "&y2=" & request.querystring("y2") & "&runNo=" & request.querystring("runNo")& "&m=" & request.querystring("m") & "&menu=" & request.querystring("menu") & "&bp=" & server.urlencode(bp))
                url = "http://" & request.serverVariables("SERVER_NAME") & "/" &  SYSTEM_FOLDER&"/"&  FIN_SUB_SYSTEM_FOLDER &"/clr/viewForm.asp?no="& request.querystring("no")& "&orgNo="& request.querystring("orgNo")& "&y="& request.querystring("y") & "&empNo=" & request.querystring("empNo") & "&y2=" & request.querystring("y2") & "&runNo=" & request.querystring("runNo")& "&m=" & request.querystring("m") & "&menu=" & request.querystring("menu") & "&bp=" & server.urlencode(bp)
			else
				'response.redirect ("http://amecweb.mitsubishielevatorasia.co.th/webflow/finform/clr/Clr/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/2" )
                url = "http://amecweb.mitsubishielevatorasia.co.th/webflow/finform/clr/Clr/index/" & session("uid") & "/" & request.querystring("no") & "/" & request.querystring("orgNo") & "/"& request.querystring("y") & "/"& request.querystring("y2") & "/"& request.querystring("runNo") &"/2"
			end if
		end if
        Call RedirectTo(url)
%>