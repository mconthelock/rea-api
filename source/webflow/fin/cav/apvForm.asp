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
<!--#include file="../../inc/connAS400.inc"-->
<%
Dim rsOra, rsOra1, sqlOra, rsAs400
Dim key, cntLine
Dim myUpload
Dim path
Dim indx, multiPart, element, flowst
Dim paydate, pfrmtime, ptotime, chkClrFrm
Dim msg
Dim fileName, name, strNewFileName
Dim amt

Set rsOra = createObject("adodb.recordset")
Set rsOra1 = createObject("adodb.recordset")
Set rsAs400 = createObject("adodb.recordset")
Set myUpload =  server.createObject("SoftArtisans.FileUp")

'Response.CharSet = "windows-874" 
cntLine = 1
chkClrFrm = false
connOra.beginTrans

select case lCase(myUpload.form("act"))
	case "reject"
		Call reject(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
	case "approve"
'		Call approve(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
		If myUpload.form("stepExt") = "01" then
			paydate = myUpload.form("SelDate") & "/" & myUpload.form("SelMon") & "/" &  myUpload.form("SelYear")		
			pfrmtime = chkLenth(myUpload.form("selBTimein")) & chkLenth(myUpload.form("selETimein"))
			ptotime  = chkLenth(myUpload.form("selBegOut")) & chkLenth(myUpload.form("selEndOut"))

			sqlOra = "update cashAdvForm set cpayby ='" & myUpload.form("radcashby") & "', dpaydate = to_date('" & paydate & "','dd/mm/yyyy'), cpayfrtime='" & pfrmtime & "', cpaytotime='" & ptotime & "'" _
			& " , ccashtype = '" & myUpload.form("radCashTy") & "'" _
			& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '"&  myUpload.form("orgNo") & "' and cyear= '" & myUpload.form("y") & "'" _
			& " and cyear2 ='" & myUpload.form("y2") & "' and nrunno = '" & myUpload.form("runNo") & "'"
'			response.write sqlOra
'			response.end
			rsOra.open sqlOra, connOra
'		
		End If
		path = server.mapPath("..") & "\cav\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" &myUpload.form("y2") & "_" &  myUpload.form("runNo") & "\"			
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
	
						sqlOra = " insert into attCashFrm (nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile)  values('" _
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
      
	'	If myUpload.form("stepExt") = "01" then		
			'#######	Insert data to datalibi.caf001		###################				
	'		sqlOra = "select crecby from cashAdvForm where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '"&  myUpload.form("orgNo") & "' 'and cyear= '" & myUpload.form("y") & "'" _
	'		& " and cyear2 ='" & myUpload.form("y2") & "' and nrunno = '" & myUpload.form("runNo") & "' "
			'response.write sqlOra
	'		rsOra.open sqlOra, connOra

	'		if not rsOra.EoF then
	'			sqlOra = "insert into datalibi.caf001(cfyer, cfrun, cfpda, cfrev, cfpay, cfsts) values('" & myUpload.form("SelYear") & "','" & myUpload.form("runNo") & "'" _
	'			& ",'" & myUpload.form("SelYear") & chkLenth(myUpload.form("SelMon")) & chkLenth(myUpload.form("SelDate")) & "','" & rsOra("crecby") & "','" & myUpload.form("radcashby") & "','N') "
				
	'			 rsAs400.open sqlOra, conn400
				'response.write sqlOra
				'		response.end
				'conn400.Close
'			'#######	Insert data to datalibi.cal001		###################				
	'			sqlOra = "select to_number(id) id, description, to_number(namount) namount from cashAdvList where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '"&  myUpload.form("orgNo") & "' and cyear= '" & myUpload.form("y") & "'" _
	'			& " and cyear2 ='" & myUpload.form("y2") & "' and nrunno = '" & myUpload.form("runNo") & "'"
	'			rsOra1.open sqlOra, connOra

	'			if not rsOra1.EoF then
	'				do while not rsOra1.EoF 
	'					sqlOra = "insert into datalibi.cal001(clyer, clrun, clseq, cldes, clamt) values('" & myUpload.form("SelYear") & "','" & myUpload.form("runNo") & "'" _
	'					& ",'" & rsOra1("id") & "','" & left(oraFormatStr(rsOra1("description")),50) & "','" & rsOra1("namount")& "') "
						'response.write sqlOra
						'response.write sqlOra
						'response.end	
	'					rsAs400.open sqlOra, conn400
						'response.write "tttttt"
							'response.end
	'					rsOra1.MoveNext
	'				loop
	'				conn400.Close
	'			end if
	'			rsOra1.Close
'			'#######	End insert data to datalibi.cal001		###################			
	'		end if
'			rsOra.Close

'			'#######	End insert data to datalibi.caf001		###################			
'		end if
		'response.end
	
		Call approve(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
		
	case "request"
		if chkWorkingUser(myUpload.form("txtReq")) then
			response.write "Requester not found!"
			response.end
		end if
		amt = 0
		for indx = 1 to myUpload.form("hdnMaxLine")
			if oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) <> "" then
					amt = amt + CDbl(trim(myUpload.form("Column2_" & indx)))
			end if
		next
		if  myUpload.form("radRecBy") <> "2" then
		   ' if myUpload.form("txtReq") <> "12225" then
				if amt > 50000 then
					 response.write "Amount over 50,000"
					 response.end
				end if
			'end if
		end if
		'Not urgent
		if myUpload.form("radUrgent") = 0 then
'			sqlOra = "select count(c.nrunno) cntform from cashAdvForm c, form" _
			sqlOra = "select form.cyear2, form.nrunno from cashAdvForm c, form" _
			& " where c.nfrmno = form.nfrmno and c.vorgno = form.vorgno and c.cyear = form.cyear " _
			& " and c.cyear2 = form.cyear2 and c.nrunno = form.nrunno and form.cst ='" & FLOW_RUNNING & "'" _
			& " and form.vreqno = '" & session("uid") & "' and " & session("divCode")  &  "<> '040101'  and " & session("divCode")  &  "<> '020101' "
			rsOra.open sqlOra, connOra
			'response.write sqlOra
		'if rsOra("cntform") > 0 then
		'	if  not rsOra.EOF and session("uid") <> "94064"  then	'มีฟอร์ม Cash Advance ค้าง
		    if  not rsOra.EOF   then
				rsOra.Close
				response.write "You can't request, please wait until old form is finished."
				response.end
			else 'ไม่มีฟอร์ม Cash Advance ค้าง
				rsOra.Close
				
				if (session("divCode") <> "040101") and (session("divCode") <> "020101") then
                           
					sqlOra = "select form.cyear2, form.nrunno from cashAdvForm c, form" _
					& " where c.nfrmno = form.nfrmno and c.vorgno = form.vorgno and c.cyear = form.cyear " _
					& " and c.cyear2 = form.cyear2 and c.nrunno = form.nrunno and form.cst ='" & FLOW_APPROVE & "'" _
					& " and form.vreqno = '" & session("uid") & "'"
					rsOra.open sqlOra, connOra

					Do while not rsOra.EoF
						sqlOra = "select * from clrForm where cyadv = '" & rsOra("cyear2") & "' and nnoadv = '" & rsOra("nrunno") & "'"
						response.write sqlOra
						rsOra1.open sqlOra, connOra
						if rsOra1.EoF then
							chkClrFrm = true
							exit do
						end if
						rsOra.MoveNext
						rsOra1.Close
					Loop
					rsOra.Close
				end if
				'if chkClrFrm = true and session("uid") <> "94064"  then
				
				if chkClrFrm = true then
					response.write "You can't request, please Clear Advance old form."
					response.end				
				end if
				
			end if
		end if
		key = createFlow(connOra,myUpload.form("frmNo"),myUpload.form("orgNo"),myUpload.form("y"),myUpload.form("txtReq"),session("uid"),myUpload.form("txtRemark"))
		if key <> false then
			key = split(key, ",")
		else
			response.write "Create request failure!"
			response.end
		end if
		if session("depCode") = "020601" then
					sqlOra = "update flow set cstepnextno = '02' where nfrmno = '"&myUpload.form("frmNo")&"' and vorgno = '"&myUpload.form("orgNo")&"' and cyear = '"&myUpload.form("y")&"' and cyear2='"&key(0)&"' and nrunno = '"&key(1)&"' and cstepnextno = '19'"
					rsOra.open sqlOra, connOra
					sqlOra = "update flow set cstepst = '1' where nfrmno = '"&myUpload.form("frmNo")&"' and vorgno = '"&myUpload.form("orgNo")&"' and cyear = '"&myUpload.form("y")&"' and cyear2='"&key(0)&"' and nrunno = '"&key(1)&"' and cstepno = '19'"
					rsOra.open sqlOra, connOra
					sqlOra = "insert into flow values('"&myUpload.form("frmNo")&"','"&myUpload.form("orgNo")&"','"&myUpload.form("y")&"','"&key(0)&"','"&key(1)&"','02','19','0','2','1','10','J0158','J0158','','0','','','','1','','2','/form/fin/cav/index.asp','','')"
					rsOra.open sqlOra, connOra	
		end if
		
'		rsOra.close
		'if myUpload.form("radRecBy") = "2" then
		'	sqlOra = "update flow set vapvno = '14217' , vrepno = '14217' where nfrmno = '"&myUpload.form("frmNo")&"' and vorgno = '"&myUpload.form("orgNo")&"' and cyear = '"&myUpload.form("y")&"' and cyear2 ='"&key(0)&"' and nrunno ='"&key(1)&"' and cextdata = '01'"
		    'response.write sqlOra
			'response.end
		'	rsOra.open sqlOra, connOra
		'end if
		sqlOra = "insert into cashAdvForm (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, sprno, curgent, crecby, effdate) " _
		& "values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "' " _
		& ",'" & myUpload.form("txtprno") & "','" & myUpload.form("radUrgent") & "','" & myUpload.form("radRecBy") & "', to_date('" & myUpload.form("SelFLDate") & "/" & myUpload.form("SelFLMon") & "/" & myUpload.form("SelFLYear") &"','dd/mm/yyyy'))"
		response.write sqlOra & "<br>"
'		response.end
		rsOra.open sqlOra, connOra
		for indx = 1 to myUpload.form("hdnMaxLine")
			if oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) <> "" then
'				sqlOra = "insert into cashAdvList (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, id, description, nquality, nprice, nvat) " _
'				& " values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "', '" & cntLine & "'" _
'				& ", '" & oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) & "', '" & oraFormatStr(trim(myUpload.form("Column2_" & indx))) & "'" _
'				&",'" & oraFormatStr(trim(myUpload.form("Column3_" & indx))) & "','" & uCase(myUpload.form("Column4_" & indx)) & "')"

				sqlOra = "insert into cashAdvList (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, id, description, namount) " _
				& " values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "', '" & cntLine & "'" _
				& ", '" & oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) & "', '" & oraFormatStr(trim(myUpload.form("Column2_" & indx))) & "')"
				
'				response.write sqlOra & "<br>"
				rsOra.open sqlOra, connOra	
				cntLine = cntLine + 1
			end if
		next
		
		cntLine = 0
'	If Not myUpload.Form("uploadFile0").IsEmpty Then
		path = server.mapPath("..") & "\cav\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" & key(0) & "_" &  key(1) & "\"			
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
	
							sqlOra = " insert into attCashFrm(nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile) values('" _
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
'	sqlOra = "select nvl2(sum((to_number(nquality)* to_number(nprice))* to_number(nvat)),sum((to_number(nquality)* to_number(nprice))* to_number(nvat)),0) as sumcash" _
'	& " from (select nquality, nprice,case nvat when 0 then 1 else (nvat/100) end as nvat from cashAdvList" _
'	& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
'	& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "')"

'jeab	sqlOra = "select sum(to_number(namount)) sumcash" _
'jeab	& " from cashAdvList" _
'jeab	& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
'jeab	& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "'"
'	response.write sqlOra
'jeab	rsOra.open sqlOra, connOra
'		response.write rsOra("sumcash")
'jeab	if 	rsOra("sumcash") <= 10000 then
		'connOra.commitTrans
'jeab		rsOra.close
'jeab		sqlOra = "select cstepnextno from flow " _
'jeab		& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
'jeab		& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and vapvno = '" & session("uid") & "'"
'jeab		rsOra.open sqlOra, connOra
'jeab	    if Not rsOra.EOF then
'jeab		if 	(rsOra("cstepnextno") <> "02") and  (rsOra("cstepnextno") <> "03") then
'jeab			rsOra.close
'jeab			sqlOra = "delete from flow where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
'jeab			& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and cstepno in ('03','02')"
'jeab			rsOra.open sqlOra, connOra			
			
'jeab			sqlOra = "update flow set cstepnextno = '19' where " _
'jeab			& " nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
'jeab			& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and cstepno ='04'"
'jeab			rsOra.open sqlOra, connOra
			
'jeab			sqlOra = "update flow set cstepst='2' where " _
'jeab			& " nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
'jeab			& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and cstepno ='19'"
'jeab			rsOra.open sqlOra, connOra
			
'jeab		end if
'jeab		end if
'jeab	else
'jeab		rsOra.close
'jeab	end if		
'	end if
'		response.end
	case else
		response.write "Invalid action!"
		response.end
end select

'response.end
flowst = getFlowStatus(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"))
if flowst = FLOW_APPROVE  then
	msg = "approve finished" & vbcrlf 
	msg = msg & "Pay Date: " & paydate & vbcrlf 
	msg = msg & "Time: " & left(pfrmtime,2) & ":" & right(pfrmtime,2) & " To " & left(ptotime,2) & ":" & right(ptotime,2)
	call sendMailRequester(myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), myUpload.form("requester"), msg)
end if
connOra.commitTrans
		
'response.end		
if (myUpload.form("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/create.asp"
else
	response.redirect (myUpload.form("bp"))
end if


%>
