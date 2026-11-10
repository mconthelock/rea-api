<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include file="../../../inc/adovbs.inc"-->
<!--#include file="../../../inc/db.inc"-->
<!--#include file="../../../inc/connUtil.inc"-->
<!--#include file="../../../inc/flowUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->
<!--#include file="../../../inc/date.inc"-->
<!--#include file="../../../inc/oraUtil.inc"-->
<!--#include file="../../../inc/authen.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!--#include virtual="/form/is/wwwRegFrm/inc/insWWW.inc"-->

<%
Dim rsOra, rsOra2, sqlOra
Dim key, i
Dim req, app, extData
Dim apv, apvList, remark

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

connOra.beginTrans

select case lCase(request.form("act"))
	case "approve"

		approve connOra, request.form("no"), request.form("orgNo"), request.form("y"), request.form("y2"), request.form("runNo"), request.form("apv"), request.form("txtRemark")
		sqlOra = "update pdmform set vpassword = '" & request.form("pass") &"', vuserno ='" & request.form("user") &"' where nfrmno = '" & request.form("no") &"' and vorgno='" & request.form("orgNo") &"' and cyear='" & request.form("y") &"' and cyear2='" & request.form("y2") &"' and nrunno='" & request.form("runNo") &"'"
		rsOra.open sqlOra,connOra
		
	case "reject"
		
		reject connOra, request.form("no"), request.form("orgNo"), request.form("y"), request.form("y2"), request.form("runNo"), request.form("apv"), request.form("txtRemark")
		
	case "request"
			
		validateData
		
		key	 = createFlow(connOra, request.form("no"), request.form("orgNo"), request.form("y"), request.form("txtReq"), request.form("apv"), request.form("txtRemark"))
		
		if key <> false then
			key = split(key, ",")
		else
			response.write "Creating request is fail!"
			response.end
		end if
		
		'########## SAVE DARWIN DATA ##########
		'Read approver for ddFlow
		sqlOra = "select distinct vApvNo from flow where nFrmNo = '" & request.form("no") & "' " _
		& "and vOrgNo = '" & request.form("orgNo") & "' and cYear = '" & request.form("y") & "' " _
		& "and cYear2 = '" & key(0) & "' and nRunNo = '" & key(1) & "'"
		rsOra.open sqlOra, connOra
		
		if not rsOra.EOF then
			apvList = rsOra("vApvNo")
			rsOra.moveNext
			
			do while not rsOra.EOF
				apvList = apvList & "," & rsOra("vApvNo")
				rsOra.moveNext
			loop
		end if
		
		rsOra.close
		
		'Save application data
		for i = 1 to 20
			if request.form("optReq" & i) <> "" then
				app = split(request.form("optReq" & i), ",")
	
				sqlOra = "insert into pdmform " _
				& "(nFrmNo, vOrgNo, cYear, cYear2, nRunNo, cNo, cReq, vPurpose) " _
				& "values " _
				& "('" & request.form("no") & "', '" & request.form("orgNo") & "', '" & request.form("y") & "', '" & key(0) & "', '" & key(1) & "' " _
				& ",  '" & request.form("Authority") & "', '" & app(1) & "', '" & request.form("Purpose") & "')"
			
				rsOra.open sqlOra, connOra
				
			end if
		next
		
		'########## END SAVE EMAIL DATA ##########
		
	case "save change"
		if request.form("txtUid") <> "" and request.form("txtPwd") <> "" then
			savePassword connOra, request.form("no"), request.form("orgNo"), request.form("y"), request.form("y2"), request.form("runNo"), oraFormatStr(left(request.form("txtUid"), 30)), oraFormatStr(left(request.form("txtPwd"), 30))
		else
			response.write "You forget entering his/her password or user id!"
			response.end
		end if
	case else
		response.write "Invalid action!"
		response.end
end select

connOra.commitTrans

if request("bp") = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & SUB_SYSTEM_FOLDER & "/create.asp"
else
	response.redirect request("bp")
end if

'########## FUNCTION DECLARATION ##########

function validateData()
	Dim requireData

	requireData = false
	
	if request.form("txtReq") = "" then
	%>
	<script language="JavaScript">
		alert("Requester not found!");
	</script>
	<%
	end if
	
	if request.form("optReq1") <> "" then
		requireData = true 
	else
	%>	
	<script language ="JavaScript">
		alert("Application requisition not found!");
	</script>
	<%
	end if 

end function

function updateDDFlow()
	Dim sqlOra, rsOra
	Dim i, app
	
	set rsOra = createObject("adodb.recordset")
	
	for i = 1 to request.form("totalApp")
		if request.form("optApprove" & i) <> "" then
			app = split(request.form("optApprove" & i), ",")
			
			sqlOra = "update klgFlow set " _
			& "cApprove = '" & app(1) & "', vRemark = '" & oraFormatStr(left(trim(request.form("txtRemark" & i)), 4000)) & "' " _
			& "where " _
			& "nFrmNo = '" & request.form("no") & "' and vOrgNo = '" & request.form("orgNo") & "' and cYear = '" & request.form("y") & "' and nRunNo = '" & request.form("runNo") & "' and cYear2 = '" & request.form("y2") & "' " _
			& " and cNo = '" & app(0) & "' " _
			& " and vApvNo = '" & request.form("apv") & "'"
		'	response.write sqlOra & "<br>"
		'	response.end
			rsOra.open sqlOra, connOra
		end if
	next
end function


'########## END FUNCTION DECLARATION ##########
%>