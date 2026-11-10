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
Dim rsOra, rsOra1, sqlOra
Dim key, cntLine
Dim myUpload
Dim path
Dim indx, multiPart, element, flowst
Dim paydate, pfrmtime, ptotime, chkClrFrm
Dim msg
Dim fileName, name, strNewFileName

Set rsOra = createObject("adodb.recordset")
Set rsOra1 = createObject("adodb.recordset")
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
	
						sqlOra = " insert into attClrFrm (nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile)  values('" _
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
		key = createFlow(connOra,myUpload.form("frmNo"),myUpload.form("orgNo"),myUpload.form("y"),myUpload.form("txtReq"),session("uid"),myUpload.form("txtRemark"))
		if key <> false then
			key = split(key, ",")
		else
			response.write "Create request failure!"
			response.end
		end if
rsOra.close
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
	
							sqlOra = " insert into attClrFrm(nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile) values('" _
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

	sqlOra = "select sum(to_number(namount)) sumcash" _
	& " from cashAdvList" _
	& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
	& " and cyear2 = '" &myUpload.form("txtyadv") & "' and nrunno= '" & myUpload.form("txtadvno")& "'"
		'response.write sqlOra
	
	rsOra.open sqlOra, connOra
'		response.write rsOra("sumcash")
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
