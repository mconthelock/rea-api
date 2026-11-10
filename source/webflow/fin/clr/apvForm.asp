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
Dim rsOra, rsOra1, sqlOra , rsAs400, rsOra2
Dim key, cntLine
Dim myUpload
Dim path
Dim indx, multiPart, element, flowst
Dim paydate, pfrmtime, ptotime, chkClrFrm
Dim msg
Dim fileName, name, strNewFileName
Dim appdt, acc, vat, amt, vat1
Dim maxid, clf

Set rsOra = createObject("adodb.recordset")
Set rsOra1 = createObject("adodb.recordset")
Set rsOra2 = createObject("adodb.recordset")
Set rsAs400 = createObject("adodb.recordset")
Set myUpload =  server.createObject("SoftArtisans.FileUp")

'Response.CharSet = "windows-874" 
cntLine = 0
chkClrFrm = false
connOra.beginTrans
select case lCase(myUpload.form("act"))
	case "reject"
		Call reject(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
	case "approve"
		path = server.mapPath("..") & "\clr\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" &myUpload.form("y2") & "_" &  myUpload.form("runNo") & "\"			
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
	
						sqlOra = " insert into attClrFrm (nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile , typeno)  values('" _
						& myUpload.form("frmNo") & "','" & myUpload.form("orgNo") & "','" & myUpload.form("y") & "','" & myUpload.form("y2") & "','" & myUpload.form("runNo") _
						& "','" & cntLine+1 & "','" & oraFormatStr(fileName) & "' , '"& mid(multiPart,11,1) &"' )"
						'response.write sqlOra & "<br>"
						'response.end
						rsOra.open sqlOra, connOra	
						cntLine = cntLine + 1
					end if 
				end if
			next
		next
		Call approve(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
		sqlOra = "select cst from form where nfrmno='"&myUpload.form("frmNo")&"' and vorgno = '"&myUpload.form("orgNo")&"' and cyear = '"&myUpload.form("y")&"' and cyear2 = '"&myUpload.form("y2")&"' and nrunno = '"&myUpload.form("runNo")&"'"	
		rsOra.open sqlOra , connOra
		if rsOra("cst") = FLOW_APPROVE then
				sqlOra = "select id, svendinv" _
							& " from clrList" _
							& " where clrList.nfrmno = '" & myUpload.form("frmNo") & "'" _
							& " and clrList.vorgno = '" & myUpload.form("orgNo")  & "'" _
							& " and clrList.cyear = '" & myUpload.form("y") & "'" _
							& " and clrList.cyear2 = '" & myUpload.form("y2")  & "'" _
							& " and clrList.nrunno = '" & myUpload.form("runNo") & "' "
				rsOra1.open sqlOra, connOra
				if not rsOra1.EoF then
					do while not rsOra1.EoF 
						acc = "acc"&rsOra1("id")
						vat = "vat"&rsOra1("id")
						amt = "amt"&rsOra1("id")
						'response.write amt
						sqlOra = "update clrList set namount = to_number("&myUpload.form(amt) &")" _
								& " where clrList.id = '"&rsOra1("id")&"' and clrList.nfrmno = '" & myUpload.form("frmNo") & "'" _
								& " and clrList.vorgno = '" & myUpload.form("orgNo")  & "'" _
								& " and clrList.cyear = '" & myUpload.form("y") & "'" _
								& " and clrList.cyear2 = '" & myUpload.form("y2")  & "'" _
								& " and clrList.nrunno = '" & myUpload.form("runNo") & "' "
						'response.write sqlOra 
						rsOra2.open sqlOra , connOra
						if myUpload.form(vat) = "" then
							vat1 = 0 
						else
							vat1 = myUpload.form(vat)
						end if
						if myUpload.form(amt) <> 0 then
							sqlOra = "insert into datalibi.cll100(cllyer,cllrun,cllseq,cllslp,cllacc,cllamt,cllvat) values('"&year(date)&"','"&myUpload.form("runNo")&"','" _
									& rsOra1("id")&"','"&Left(rsOra1("svendinv"),30)&"','"&myUpload.form(acc)&"','"&myUpload.form(amt)&"','"&vat1&"')"
							'response.write sqlOra
							'rsAs400.open sqlOra, conn400
						end if
						maxid = CInt(rsOra1("id"))
						rsOra1.MoveNext
					loop
					'conn400.Close
				end if
				rsOra1.Close
				if myUpload.form("htax") <> "" and myUpload.form("htax") <> "0" then
					maxid = maxid+1
					sqlOra = "insert into datalibi.cll100(cllyer,cllrun,cllseq,cllslp,cllacc,cllamt,cllvat) values('"&year(date)&"','"&myUpload.form("runNo")&"','" _
									& maxid &"','withholding tax','2143004','"&myUpload.form("htax")&"','0')"
					'rsAs400.open sqlOra, conn400
				end if
				
				sqlOra = "select sum(namount) clfamt" _	
							& " from clrList" _
							& " where clrList.nfrmno = '" & myUpload.form("frmNo") & "'" _
							& " and clrList.vorgno = '" & myUpload.form("orgNo")  & "'" _
							& " and clrList.cyear = '" & myUpload.form("y") & "'" _
							& " and clrList.cyear2 = '" & myUpload.form("y2")  & "'" _
							& " and clrList.nrunno = '" & myUpload.form("runNo") & "' group by clrList.nrunno"
				rsOra1.open sqlOra, connOra
				if myUpload.form("htax") <> "" and myUpload.form("htax") <> "0" then
					clf = CDbl(rsOra1("clfamt")) - CDbl(myUpload.form("htax"))
					sqlOra = "update clrForm set HOLDTAX = to_number("&myUpload.form("htax") &") " _
							& " where nfrmno = '" & myUpload.form("frmNo") & "'" _
							& " and vorgno = '" & myUpload.form("orgNo")  & "'" _
							& " and cyear = '" & myUpload.form("y") & "'" _
							& " and cyear2 = '" & myUpload.form("y2")  & "'" _
							& " and nrunno = '" & myUpload.form("runNo") & "'"
					rsOra2.open sqlOra , connOra
					
				else
					clf = rsOra1("clfamt")
				end if			
				
				'sqlOra = "insert into datalibi.clf100(clfyer, clfrun, clfdat, clfamt, clfadv, clfctr,clfdes) values('" & myUpload.form("y2") & "','" & myUpload.form("runNo") & "'" _
				sqlOra = "insert into datalibi.clf100(clfyer, clfrun, clfdat, clfamt, clfadv, clfctr,clfdes) values('" & year(date) & "','" & myUpload.form("runNo") & "'" _
				& ",'" &year(date)&chkLenth(month(date))& chkLenth(day(date)) & "','" & clf & "','" & myUpload.form("clfadv") & "','"&myUpload.form("costcn")&"','"&left(oraFormatStr(myUpload.form("clfdes")),30)&"')"
				'response.write sqlOra
				'response.end
				'rsAs400.open sqlOra, conn400
				conn400.Close	
		end if
		rsOra.close
		
		
	case "request"
		if chkWorkingUser(myUpload.form("txtReq")) then
			response.write "Requester not found!"
			response.end
		end if
		sqlOra = "select * from cashAdvForm cav, form where cav.cyear2 = '" & myUpload.form("txtyadv") & "'" _
		& " and cav.nrunno = '" & myUpload.form("txtadvno") & "' and form.vreqno = '" & myUpload.form("txtReq") & "'" _
		& " and cav.nfrmno = form.nfrmno and cav.vorgno = form.vorgno and cav.cyear = form.cyear " _
		& " and cav.cyear2 = form.cyear2 and cav.nrunno = form.nrunno  and form.cst = '" & FLOW_APPROVE & "'" 
		'response.write sqlOra
		rsOra.open sqlOra, connOra
		if rsOra.EoF then
			response.write "Please Clear Advance after Cash Advance Form Finished."
			response.end
		end if
		rsOra.close
		sqlOra = "select clrForm.* from clrForm, form " _
		 & " where cyadv = '" & myUpload.form("txtyadv") & "' and nnoadv = '" & myUpload.form("txtadvno") & "'" _
		 & " and clrForm.nfrmno = form.nfrmno and clrForm.vorgno = form.vorgno and clrForm.cyear = form.cyear and clrForm.cyear2 = form.cyear2 and clrForm.nrunno = form.nrunno " _
		 & " and form.cst in ('" & FLOW_RUNNING & "','" & FLOW_APPROVE & "')"
		 'response.write sqlOra
		 
		rsOra.open sqlOra, connOra
		if not rsOra.EoF then
			response.write "Request Clear Form is repeat. You can't request form." 
			response.end
		end if
		rsOra.close
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

		sqlOra = "insert into clrForm (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, cyadv, nnoadv) " _
		& "values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "' " _
		& ",'" & myUpload.form("txtyadv") & "','" & myUpload.form("txtadvno") & "')"
'		response.write sqlOra & "<br>"
'		response.end
		rsOra.open sqlOra, connOra
		
		for indx = 1 to myUpload.form("hdnMaxLine")
			if oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) <> "" then
				sqlOra = "insert into clrList (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, id, svendinv, namount) " _
				& " values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "', '" & cntLine + 1 & "'" _
				& ", '" & oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) & "', '" & oraFormatStr(trim(myUpload.form("Column2_" & indx))) & "')"				
'				response.write sqlOra & "<br>"
				rsOra.open sqlOra, connOra	
				cntLine = cntLine + 1
			end if
		next
		if cntLine = 0 then
			response.write "Please input detail of Receipt No"
			response.end
		end if 
		cntLine = 0
'	If Not myUpload.Form("uploadFile0").IsEmpty Then
		path = server.mapPath("..") & "\clr\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" & key(0) & "_" &  key(1) & "\"			
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
	
							sqlOra = " insert into attClrFrm(nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile , typeno ) values('" _
							& myUpload.form("frmNo") & "','" & myUpload.form("orgNo") & "','" & myUpload.form("y") & "','" & key(0) & "','" &  key(1) _
							& "','" & cntLine+1 & "','" & oraFormatStr(fileName) & "' , '"& mid(multiPart,11,1) &"' )"
							'response.write sqlOra & "<br>"
							'response.end
							rsOra.open sqlOra, connOra	
							cntLine = cntLine + 1
					end if 
				end if
			next
		next				

	sqlOra = "select sum(to_number(namount)) sumcash" _
	& " from cashAdvList" _
	& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
	& " and cyear2 = '" &myUpload.form("txtyadv") & "' and nrunno= '" & myUpload.form("txtadvno")& "'"
		'response.write sqlOra
	
	rsOra.open sqlOra, connOra

	if 	rsOra("sumcash") <= 10000 then
		rsOra.close
		sqlOra = "delete from flow where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
		& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and cstepno in ('03','02')"
		rsOra.open sqlOra, connOra			
		
		sqlOra = "update flow set cstepnextno = '19' where " _
		& " nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
		& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "' and cstepno ='04'"
		rsOra.open sqlOra, connOra
	else
		rsOra.close
	end if		

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
connOra.commitTrans
		

if (myUpload.form("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/create.asp"
else
	response.redirect (myUpload.form("bp"))
end if


%>
