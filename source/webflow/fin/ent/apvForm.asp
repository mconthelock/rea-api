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
Dim purpose, description
Dim msg
Dim fileName, name, strNewFileName

Set rsOra = createObject("adodb.recordset")
Set rsOra1 = createObject("adodb.recordset")
Set myUpload =  server.createObject("SoftArtisans.FileUp")

'Response.CharSet = "windows-874" 
cntLine = 1
chkClrFrm = false
connOra.beginTrans
'response.write myUpload.form("act")
select case lCase(myUpload.form("act"))
	case "reject"
		Call reject(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
	case "approve"
		' path = server.mapPath("..") & "\ent\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" &myUpload.form("y2") & "_" &  myUpload.form("runNo") & "\"			
		' createOSDir path	
		' myUpload.path = path
		' for each multiPart in myUpload.Form
			' for each element in myUpload.FormEx(multiPart)
				' if IsObject( element ) then
					' if not ( element.IsEmpty ) then
						' fileName = mid(myUpload.Form(multiPart).UserFilename,instrrev(myUpload.Form(multiPart).UserFilename,"\")+1)
						' fileName = name &"" & fileName
						' 'fileName = replace(replace(fileName," ",""),"'","&#39;")
						' fileName = replace(replace(fileName," ",""),  "'", "-")
						' strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)
						' fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName

						' element.SaveAs path& filename
	
						' sqlOra = " insert into attEntFrm (nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile)  values('" _
						' & myUpload.form("frmNo") & "','" & myUpload.form("orgNo") & "','" & myUpload.form("y") & "','" & myUpload.form("y2") & "','" & myUpload.form("runNo") _
						' & "','" & cntLine+1 & "','" & oraFormatStr(fileName) & "')"
						' 'response.write sqlOra & "<br>"
						' 'response.end
						' rsOra.open sqlOra, connOra	
						' cntLine = cntLine + 1
					' end if 
				' end if
			' next
		' next

		Call approve(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
		
	case "request"
		if chkWorkingUser(myUpload.form("txtReq")) then
			response.write "Requester not found!"
			response.end
		end if
		key = createFlow(connOra,myUpload.form("frmNo"),myUpload.form("orgNo"),myUpload.form("y"),myUpload.form("txtReq"),session("uid"),myUpload.form("txtRemark"))
		if key <> false then
			key = split(key, ",")
			response.write key(1)
		else
			response.write "Create request failure!"
			response.end
		end if
		
		'Purpoose
		if myUpload.form("typePurpose") = "OTHERS" then
			purpose = myUpload.form("txtOtherPurpose") 
		else
			purpose = myUpload.form("typePurpose") 
		end if
		
		'Description
		if myUpload.form("typeDescript") = "OTHERS" then
			description = myUpload.form("txtOtherDescript") 
		else 
			description = myUpload.form("typeDescript") 
		end if
		
		sqlOra = "insert into rqtEntForm (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, payto, purpose, paydate, place, description) " _
		& "values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "' " _
		& ",'" & myUpload.form("txtReq") & "','" & purpose & "', to_date('" & myUpload.form("SelFLDate") & "/" & myUpload.form("SelFLMon") & "/" & myUpload.form("SelFLYear") & "','dd/mm/yyyy'),'" & myUpload.form("txtPlace") & "' ,'" & description & "')"
		rsOra.open sqlOra, connOra	
		
		'Expense List
		cntLine = 1
		for indx = 1 to CInt(myUpload.form("expenseMaxline"))
			sqlOra = "insert into rqtEntList (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, id, description, namount) " _
			& " values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "', '" & cntLine & "'" _
			& ", '" & myUpload.form("partExpense" & indx) & "', '" & myUpload.form("priceExpense" & indx) & "')"				
			rsOra.open sqlOra, connOra	
			cntLine = cntLine + 1
		next
		
		'Person List
		cntLine = 1
		for indx = 1 to CInt(myUpload.form("personsMaxline"))
			sqlOra = "insert into rqtEntPersons (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, id, pname, ptype) " _
			& " values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "', '" & cntLine & "'" _
			& ", '" & myUpload.form("namePersons" & indx) & "', '" & myUpload.form("typeGuest" & indx) & "')"				
			rsOra.open sqlOra, connOra	
			cntLine = cntLine + 1
		next
		
		cntLine = 1
		path = server.mapPath("..") & "\ent\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" & key(0) & "_" &  key(1) & "\"			
		createOSDir path
		myUpload.path = path
		for each multiPart in myUpload.FormEx
			if IsObject(myUpload.Form(multiPart)) then
				fileName = mid(myUpload.Form(multiPart).UserFilename,instrrev(myUpload.Form(multiPart).UserFilename,"\")+1)
				'fileName = name & "" & fileName
				fileName = replace(replace(fileName," ",""),  "'", "-")
					
			    'strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)

			    'fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName

				myUpload.Form(multiPart).SaveAs path & fileName
	
				sqlOra = " insert into attEntFrm(nfrmno, vorgno, cyear, cyear2, nrunno, id, sfile) values('" _
				& myUpload.form("frmNo") & "','" & myUpload.form("orgNo") & "','" & myUpload.form("y") & "','" & key(0) & "','" &  key(1) _
				& "','" & cntLine & "','" & oraFormatStr(fileName) & "')"
				rsOra.open sqlOra, connOra	
				cntLine = cntLine + 1
			end if
		next				

	case else
		response.write "Invalid action!"
		response.end
end select

'response.end
' flowst = getFlowStatus(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"))
' if flowst = FLOW_APPROVE  then
	' msg = "approve finished" & vbcrlf 
	' msg = msg & "Pay Date: " & paydate & vbcrlf 
	' msg = msg & "Time: " & left(pfrmtime,2) & ":" & right(pfrmtime,2) & " To " & left(ptotime,2) & ":" & right(ptotime,2)
	'call sendMailRequester(myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), myUpload.form("requester"), msg)
'end if

connOra.commitTrans
		
response.end		
if (myUpload.form("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/create.asp"
else
	response.redirect (myUpload.form("bp"))
end if


%>
