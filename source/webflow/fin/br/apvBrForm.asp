<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/oraUtil.inc"-->
<!--#include virtual="/form/inc/date.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!--#include virtual="/form/inc/flowUtil.inc"-->
<!--#include virtual="/form/inc/dirUtil.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->

<!--#include file="inc/flowUtil.inc"-->

<%
Dim sqlOra , rsOra , rsOra2
Dim cntLine , totassetNo , assetNo , assetNo1
Dim ssec , sdep , sdiv
Dim chk , key
Dim sreceiver , owner , receiver
Dim FLOCCODE , TLOCCODE
Dim borrowto , Empcode , borrowtype
set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

connOra.beginTrans
select case lCase(request("act"))
	case "reject"
		 Call reject(connOra, request("frmNo"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid"), request("txtRemark"))
	case "approve"
		Call approve(connOra, request("frmNo"), request("orgNo"), request("y"), request("y2"), request("runNo"), session("uid"), request("txtRemark"))	
		if request("stepExt") = "02" then
				sqlOra = "select vapvno , vrepno from flow where nfrmno='"&request("frmNo")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cextdata = '04'"
				rsOra2.open sqlOra , connOra
				if rsOra2("vapvno") = session("uid") or rsOra2("vrepno") = session("uid") then
					sqlOra = "update flow set cstepst = '2', vrealapv = '' , capvstno = '0' , dapvdate = '' , capvtime = '' , vremark = '' where nfrmno='"&request("frmNo")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"' and cextdata = '04'"
					response.write "update = "&sqlOra
					rsOra.open sqlOra , connOra
				end if
				rsOra2.close
		end if
	'	sqlOra = "select cst from form where nfrmno='"&request("frmNo")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"'"
	'	rsOra2.open sqlOra , connOra
	'	chk = 0
	'	if rsOra2("cst") = FLOW_APPROVE then
	'		chk = 1
	'	end if
	'	rsOra2.close
	'	if chk = 1 then
	'		sqlOra = "select assetno , tloccode from fxa_transForm where NFRMNO = '"&request("frmNo")&"' and vorgno = '"&request("orgNo") _
	'			&"' and cyear = '"&request("y")&"' and cyear2 = '"&request("y2")&"' and nrunno = '"&request("runNo")&"'"
	'		rsOra.open sqlOra, connOra
	'		do while not rsOra.EOF
	'			sqlOra = "update fxa_faMst set LOCCODE = '"&rsOra("tloccode")&"' where ASSETNO = '"&rsOra("assetno")&"' "
	'			rsOra2.open sqlOra, connOra
	'		rsOra.moveNext
	'		loop
	'		rsOra.close
	'	end if
	case "request"
		if chkWorkingUser(request("empReq")) then
%>
						<script language="javascript">
							alert("Requester not found");
							history.back();
						</script>
<%		
					response.end
		end if 
		'if request("actbr") = "BI" then
		'if chkWorkingUser(request("bremp")) then
%>
						<!--script language="javascript">
							alert("EMP. CODE not found");
							history.back();
						</script-->
<%		
		'			response.end
		'end if 
		cntLine = 1
		chk = 0 
		for indx = 1 to CInt(request("personsMaxline"))
			sqlOra = "select count(ASSETNO) cntrow from fxa_faMst where ASSETNO = '"&request("txtAss" & indx)&"' and statusass <> 'R'"
			rsOra.open sqlOra, connOra
			if rsOra("cntrow") = 0 then
					chk = 1
			end if
			rsOra.close
			if  indx = 1 then
				totassetNo = "'" & request("txtAss" & indx)  & "'"
				assetNo = indx & "-'" & request("txtAss" & indx)  & "'"
			else
				totassetNo = totassetNo & ",'" & request("txtAss" & indx)  & "'"			
				assetNo = assetNo & "," & indx & "-'" & request("txtAss" & indx) 	& "'"
			end if
		next 
		
		if chk = "1" then
			%>
						<script language="javascript">
							alert("AssetNo incorrect");
							history.back();
						</script>
			<%		
					response.end
		end if
		
		'response.write "ttt = "&totassetNo
		sqlOra = "select count(distinct vempno ) as cntrow from  fxa_faMst, fxa_locMst , orgpos  " _
			   &" where fxa_faMst.LOCCODE = fxa_locMst.LOCCODE and fxa_locMst.vorgno = orgpos.vorgno and fxa_locMst.sposcode = orgpos.vposno" _
			   &" and fxa_faMst.ASSETNO in (" & totassetNo & ")"
		'response.write sqlOra
		rsOra.open sqlOra, connOra
		if Not rsOra.EOF then
				'response.write "cntrow = "&rsOra("cntrow")
				if rsOra("cntrow") = "0" then
		  %>
						<script language="javascript">
							alert("AssetNo incorrect");
							history.back();
						</script>
			<%		
					response.end
				elseif rsOra("cntrow") > "1" then
		  %>
						<script language="javascript">
							alert("Many location");
							history.back();
						</script>
			<%		
					response.end				
				
					
				
				end if
		else
		
			%>
						<script language="javascript">
							alert("AssetNo not found");
							history.back();
						</script>
			<%		
					response.end
		
		end if
		rsOra.close
		
		sqlOra = "select sdivcode , sdepcode , sseccode from amec.aemployee where sempno = '"&request("empReq")&"'"
		rsOra.open sqlOra, connOra	
		ssec = rsOra("sseccode")
		sdep = rsOra("sdepcode")
		sdiv = rsOra("sdivcode")
		rsOra.close
		chk = 0
		sqlOra = "select fxa_locMst.vorgno from  fxa_faMst, fxa_locMst , orgpos  " _
			   &" where fxa_faMst.LOCCODE = fxa_locMst.LOCCODE and fxa_locMst.vorgno = orgpos.vorgno and fxa_locMst.sposcode = orgpos.vposno" _
			   &" and fxa_faMst.ASSETNO in (" & totassetNo & ")"
		'response.write "sqlOra = "&sqlOra
		rsOra.open sqlOra, connOra
		do while not rsOra.EOF
			if ssec = "00" then
				if sdep = "00" then
					if left(sdiv, 3) = left(rsOra("vorgno"),3) then
						chk = 1
					end if
				else
					if left(sdep, 4) = left(rsOra("vorgno"),4) or rsOra("vorgno") = sdiv then
						chk = 1
					end if
				end if
					
			else
				if rsOra("vorgno") = ssec or rsOra("vorgno") = sdep or rsOra("vorgno") = sdiv then
					chk = 1
				end if
			end if
			
			rsOra.moveNext
		loop
		rsOra.close
		if chk = "1" and request("actbr") = "BI"  then
			%>
						<script language="javascript">
							alert("AssetNo in your location");
							history.back();
						</script>
			<%		
					response.end
		end if
		'response.write "inter now"
		

		if request("actbr") = "BI" then
				receiver = request("empReq")
				borrowto = ""
				borrowtype = "I"
		else
				receiver = request("empReq")
				borrowto = request("brlc")
				borrowtype = "O"
		end if
		sqlOra = "select distinct vempno as owner from  fxa_faMst, fxa_locMst , orgpos  " _
			   &" where fxa_faMst.LOCCODE = fxa_locMst.LOCCODE and fxa_locMst.vorgno = orgpos.vorgno and fxa_locMst.sposcode = orgpos.vposno" _
			   &" and fxa_faMst.ASSETNO in (" & totassetNo & ")"
		'response.write sqlOra
		rsOra.open sqlOra, connOra
		owner = rsOra("owner")
		rsOra.close
		response.write owner & " " & receiver
	    key = createBrFlow(connOra,request("frmNo"),request("orgNo"),request("y"),request("empReq"),session("uid"),request("txtRemark"),owner,receiver)
		key = split(key, ",")
		' reset step admin
		sqlOra = "select vapvno , vrepno from flow where nfrmno='"&request("frmNo")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&key(0)&"' and nrunno = '"&key(1)&"' and cextdata = '04'"
		rsOra2.open sqlOra , connOra
		if rsOra2("vapvno") = request("empReq") or rsOra2("vrepno") = request("empReq") then
			sqlOra = "update flow set cstepst = '1', vrealapv = '' , capvstno = '0' , dapvdate = '' , capvtime = '' , vremark = '' where nfrmno='"&request("frmNo")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&key(0)&"' and nrunno = '"&key(1)&"' and cextdata = '04'"
			rsOra.open sqlOra , connOra
		end if
		rsOra2.close
			' reset step receive
		sqlOra = "select vapvno , vrepno from flow where nfrmno='"&request("frmNo")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&key(0)&"' and nrunno = '"&key(1)&"' and cextdata = '02'"
		rsOra2.open sqlOra , connOra
		if rsOra2("vapvno") = request("empReq") or rsOra2("vrepno") = request("empReq") then
			sqlOra = "update flow set cstepst = '2', vrealapv = '' , capvstno = '0' , dapvdate = '' , capvtime = '' , vremark = '' where nfrmno='"&request("frmNo")&"' and vorgno = '"&request("orgNo")&"' and cyear = '"&request("y")&"' and cyear2 = '"&key(0)&"' and nrunno = '"&key(1)&"' and cextdata = '02'"
			rsOra.open sqlOra , connOra
		end if
		rsOra2.close
		for indx = 1 to CInt(request("personsMaxline"))
			sqlOra = "select LOCCODE from fxa_faMst where ASSETNO='"& request("txtAss" & indx) &"'"
			rsOra2.open sqlOra, connOra
			FLOCCODE = rsOra2("LOCCODE")
			rsOra2.close
			sqlOra = "insert into fxa_borrowform values('"&request("frmNo")&"','"&request("orgNo")&"','"&request("y")&"','"&key(0)&"','"&key(1)&"','"&request("txtAss" & indx) &"'" _
					&",'"&request("txtQty" & indx)&"','"&FLOCCODE&"','','"&borrowto&"','"&oraFormatStr(left(request("txtRmNo" & indx), 1024))&"','"&receiver&"','"&borrowtype&"','"&oraFormatStr(left(request("dt"),1024))&"','"&oraFormatStr(left(request("txtcause"),1024))&"')"
			'response.write "sqlOra insert = "&sqlOra
			rsOra.open sqlOra, connOra
		next 
		'response.end
	case else
				response.write "Invalid action!"
				response.end
end select
connOra.commitTrans
if (request("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/fixAss/rtoMstf.asp"
else
	response.redirect (request("bp"))
end if


%>