<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/date.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/dirUtil.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!--#include file="../../inc/misc.inc"-->
<%
Dim rsOra, sqlOra
Dim key, cntLine
Dim myUpload
Dim path
Dim indx, multiPart, element, flowst
Dim paydate, pfrmtime, ptotime
Dim msg
Dim fileName, name, strNewFileName
Dim reqPos ,reqDep
Set rsOra = createObject("adodb.recordset")
Set myUpload =  server.createObject("SoftArtisans.FileUp")
'Response.CharSet = "windows-874" 
cntLine = 1

connOra.beginTrans
select case lCase(myUpload.form("act"))
	case "reject"
		Call reject(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
	case "approve"
		If myUpload.form("stepExt") = "01" then
			paydate = myUpload.form("SelDate") & "/" & myUpload.form("SelMon") & "/" &  myUpload.form("SelYear")		
			pfrmtime = chkLenth(myUpload.form("selBTimein")) & chkLenth(myUpload.form("selETimein"))
			ptotime  = chkLenth(myUpload.form("selBegOut")) & chkLenth(myUpload.form("selEndOut"))

'			sqlOra = "update RcptRfndForm set cpayby ='" & myUpload.form("radcashby") & "', dpaydate = to_date('" & paydate & "','dd/mm/yyyy'), cpayfrtime='" & pfrmtime & "', cpaytotime='" & ptotime & "'" _
			sqlOra = "update RcptRfndForm set dpaydate = to_date('" & paydate & "','dd/mm/yyyy'), cpayfrtime='" & pfrmtime & "', cpaytotime='" & ptotime & "'" _
			& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '"&  myUpload.form("orgNo") & "' and cyear= '" & myUpload.form("y") & "'" _
			& " and cyear2 ='" & myUpload.form("y2") & "' and nrunno = '" & myUpload.form("runNo") & "'"
'			response.write sqlOra
'			response.end
			rsOra.open sqlOra, connOra
		End If
		path = server.mapPath("..") & "\rcrf\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" &myUpload.form("y2") & "_" &  myUpload.form("runNo") & "\"			
	'	if fso.FolderExists(path) = false then
	'		response.write "if"
	'		createOSDir path
	'	else
	'		response.write "else"
	'	end if		
	'	response.end
			createOSDir path	
			myUpload.path = path	

		for each multiPart in myUpload.Form
			for each element in myUpload.FormEx(multiPart)
				if IsObject( element ) then
					if not ( element.IsEmpty ) then
						fileName = mid(myUpload.Form(multiPart).UserFilename,instrrev(myUpload.Form(multiPart).UserFilename,"\")+1)
						fileName = name &"" & fileName
						'fileName = replace(replace(fileName," ",""),"'","&#39;")
						fileName = replace(replace(fileName," ",""),  "'", "-")
						strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)
						fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName

						element.SaveAs path& filename
	
							sqlOra = " insert into attRcptRfndForm(nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile) values('" _
							& myUpload.form("frmNo") & "','" & myUpload.form("orgNo") & "','" & myUpload.form("y") & "','" & myUpload.form("y2") & "','" & myUpload.form("runNo") _
							& "','" & cntLine+1 & "','" & oraFormatStr(fileName) & "')"
							'response.write sqlOra & "<br>"
							'response.end
							rsOra.open sqlOra, connOra	
							cntLine = cntLine + 1
					end if 
				end if
			next
		next				
		Call approve(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
		
	case "request"
		if chkWorkingUser(myUpload.form("txtReq")) then
			response.write "Requester not found!"
			response.end
		end if
		key = createFlow(connOra,myUpload.form("frmNo"),myUpload.form("orgNo"),myUpload.form("y"),myUpload.form("txtReq"),session("uid"),myUpload.form("txtRemark"))

		if key <> false then
			key = split(key, ",")
			if session("depCode") = "020601" then
					sqlOra = "update flow set cstepnextno = '02' where nfrmno = '"&myUpload.form("frmNo")&"' and vorgno = '"&myUpload.form("orgNo")&"' and cyear = '"&myUpload.form("y")&"' and cyear2='"&key(0)&"' and nrunno = '"&key(1)&"' and cstepnextno = '19'"
					rsOra.open sqlOra, connOra
					sqlOra = "update flow set cstepst = '1' where nfrmno = '"&myUpload.form("frmNo")&"' and vorgno = '"&myUpload.form("orgNo")&"' and cyear = '"&myUpload.form("y")&"' and cyear2='"&key(0)&"' and nrunno = '"&key(1)&"' and cstepno = '19'"
					rsOra.open sqlOra, connOra
					sqlOra = "insert into flow values('"&myUpload.form("frmNo")&"','"&myUpload.form("orgNo")&"','"&myUpload.form("y")&"','"&key(0)&"','"&key(1)&"','02','19','0','2','1','10','J0158','J0158','','0','','','','1','','2','/form/fin/cav/index.asp','','')"
					rsOra.open sqlOra, connOra	
			end if
		else
			response.write "Create request failure!"
			response.end
		end if
		sqlOra = "insert into RcptRfndForm (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, sprno) " _
		& "values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "','" & myUpload.form("txtprno") & "')"
		'response.write sqlOra & "<br>"
		'response.end
		rsOra.open sqlOra, connOra
		
		for indx = 1 to myUpload.form("hdnMaxLine")
			if oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) <> "" then
				sqlOra = "insert into RcptRfndFormList (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, id, description, namount) " _
				& " values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "', '" & cntLine & "'" _
				& ", '" & oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) & "', '" & oraFormatStr(trim(myUpload.form("Column2_" & indx))) & "')"			
				response.write sqlOra & "<br>"
				rsOra.open sqlOra, connOra	
				cntLine = cntLine + 1
			end if
		next
		
		cntLine = 0
'	If Not myUpload.Form("uploadFile0").IsEmpty Then
		path = server.mapPath("..") & "\rcrf\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" & key(0) & "_" &  key(1) & "\"			
		createOSDir path
		myUpload.path = path
		
		for each multiPart in myUpload.Form
			for each element in myUpload.FormEx(multiPart)
				if IsObject( element ) then
					if not ( element.IsEmpty ) then
						fileName = mid(myUpload.Form(multiPart).UserFilename,instrrev(myUpload.Form(multiPart).UserFilename,"\")+1)
						fileName = name &"" & fileName
						'fileName = replace(replace(fileName," ",""),"'","&#39;")
						fileName = replace(replace(fileName," ",""),  "'", "-")
							
						strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)


						fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName

						element.SaveAs path& filename
	
							sqlOra = " insert into attRcptRfndForm(nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile) values('" _
							& myUpload.form("frmNo") & "','" & myUpload.form("orgNo") & "','" & myUpload.form("y") & "','" & key(0) & "','" &  key(1) _
							& "','" & cntLine+1 & "','" & oraFormatStr(fileName) & "')"
							'response.write sqlOra & "<br>"
							'response.end
							rsOra.open sqlOra, connOra	
							cntLine = cntLine + 1
					end if 
				end if
			next
		next				

		'sqlOra = "select sposcode, sdepcode from amec.aemployee where sempno  = '" & myUpload.form("txtReq") & "'"
		'rsOra.open sqlOra, connOra
		'reqPos = rsOra("sposcode")
		'reqDep = rsOra("sdepcode")
		'response.end
		'rsOra.Close

		'sqlOra = "select sum(to_number(namount)) sumcash" _
		'& " from RcptRfndFormList" _
		'& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
		'& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "'"
		'response.write sqlOra
		'rsOra.open sqlOra, connOra
		'if 	rsOra("sumcash") <= 10000 then
		'	rsOra.close
		'	if reqPos > "20" then
		'		if reqDep <> "00" then
		'				sqlOra = "delete from flow where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
		'				& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and cstepno in ('03','02')"
		'				rsOra.open sqlOra, connOra			
		'				
		'				sqlOra = "update flow set cstepnextno = '19' where " _
		'				& " nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
		'				& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and cstepno ='04'"
		'				rsOra.open sqlOra, connOra
		'		end if
		'	end if
		'else
		'	rsOra.close
		'end if		
	case else
		response.write "Invalid action!"
		response.end
end select

flowst = getFlowStatus(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"))
if flowst = FLOW_APPROVE  then
	msg = "approve finished" & vbcrlf 
	msg = msg & "Pay Date: " & paydate & vbcrlf 
	msg = msg & "Time: " & left(pfrmtime,2) & ":" & right(pfrmtime,2) & " To " & left(ptotime,2) & ":" & right(ptotime,2)	
	call sendMailRequester(myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), myUpload.form("requester"), msg)
end if

'response.end		
connOra.commitTrans
		
if (myUpload.form("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/create.asp"
else
	response.redirect (myUpload.form("bp"))
end if

%>
