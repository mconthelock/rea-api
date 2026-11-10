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

Set rsOra = createObject("adodb.recordset")
Set myUpload =  server.createObject("SoftArtisans.FileUp")

Response.CharSet = "windows-874" 
cntLine = 1

connOra.beginTrans

select case lCase(myUpload.form("act"))
	case "reject"
		Call reject(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
	case "approve"
		Call approve(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
		If myUpload.form("stepExt") = "01" then
			sqlOra = "update cashAdvForm set cpayby ='" & myUpload.form("radcashby") & "'" _
			& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '"&  myUpload.form("orgNo") & "' and cyear= '" & myUpload.form("y") & "'" _
			& " and cyear2 ='" & myUpload.form("y2") & "' and nrunno = '" & myUpload.form("runNo") & "'"
			'response.write sqlOra
			rsOra.open sqlOra, connOra
		End If

		
	case "request"
		if chkWorkingUser(myUpload.form("txtReq")) then
			response.write "Requester not found!"
			response.end
		end if

		key = createFlow(connOra,myUpload.form("frmNo"),myUpload.form("orgNo"),myUpload.form("y"),myUpload.form("txtReq"),session("uid"),myUpload.form("txtRemark"))
		if key <> false then
			key = split(key, ",")
		else
			response.write "Create request failure!"
			response.end
		end if

		sqlOra = "insert into cashAdvForm (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, sprno) " _
		& "values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "' " _
		& ",'" & myUpload.form("txtprno") & "')"
'		response.write sqlOra & "<br>"
'		response.end
		rsOra.open sqlOra, connOra
		
		for indx = 1 to myUpload.form("hdnMaxLine")
			if oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) <> "" then
				sqlOra = "insert into cashAdvList (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, id, description, nquality, nprice, nvat) " _
				& " values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "', '" & cntLine & "'" _
				& ", '" & oraFormatStr(left(trim(myUpload.form("Column1_" & indx)),256)) & "', '" & oraFormatStr(trim(myUpload.form("Column2_" & indx))) & "'" _
				&",'" & oraFormatStr(trim(myUpload.form("Column3_" & indx))) & "','" & uCase(myUpload.form("Column4_" & indx)) & "')"
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


						fileName = year(date) & countChar(month(date),2) & countChar(day(date),2) & countChar(Hour(Now),2) & countChar(Minute(Now),2)  & "_" & strNewFileName

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
	sqlOra = "select nvl2(sum((to_number(nquality)* to_number(nprice))* to_number(nvat)),sum((to_number(nquality)* to_number(nprice))* to_number(nvat)),0) as sumcash" _
	& " from (select nquality, nprice,case nvat when 0 then 1 else (nvat/100) end as nvat from cashAdvList" _
	& " where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
	& " and cyear2 = '" & key(0) & "' and nrunno= '" & key(1) & "')"

'		response.write sqlOra
	
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
connOra.commitTrans
		
'response.end		

'flowst = getFlowStatus(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"))
'if flowst = FLOW_APPROVE  then
'	call sendMailRequester(myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), myUpload.form("requester"), "approve finished")
'end if
'	response.end
if (myUpload.form("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/create.asp"
else
	response.redirect (myUpload.form("bp"))
end if


%>
