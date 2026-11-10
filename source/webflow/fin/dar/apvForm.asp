<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/adovbs.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="inc/flowUtil.inc"-->
<!--#include file="../../inc/date.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/dirUtil.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!--#include file="../../inc/misc.inc"-->
<%
Dim rsOra, rsOra1, sqlOra
Dim key, cntLine, apvno
Dim assetNo, assetNo1, assetNo2, totassetNo
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
assetNo = ""
connOra.beginTrans
'response.write myUpload.form("act")
select case lCase(myUpload.form("act"))
	case "reject"
		Call reject(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
	case "approve"
		Call approve(connOra, myUpload.form("frmNo"), myUpload.form("orgNo"), myUpload.form("y"), myUpload.form("y2"), myUpload.form("runNo"), session("uid"), myUpload.form("txtRemark"))	
	case "request"
		if chkWorkingUser(myUpload.form("txtReq")) then
			response.write "Requester not found!"
			response.end
		end if

		'response.write myUpload.form("selWriteOff")
		'response.end
		'sqlOra = "insert into fxa_darform (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, grpcode, loccode, assetno,nqty, vUnit,writeoffno, vremark) " _
		'& "values ('" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "' " _
		'& ",'" & myUpload.form("txtReq") & "','" & purpose & "', to_date('" & myUpload.form("SelFLDate") & "/" & myUpload.form("SelFLMon") & "/" & myUpload.form("SelFLYear") & "','dd/mm/yyyy'),'" & myUpload.form("txtPlace") & "' ,'" & description & "')"
		'rsOra.open sqlOra, connOra	

		'Person List
		cntLine = 1
		for indx = 1 to CInt(myUpload.form("personsMaxline"))
			if  indx = 1 then
				totassetNo = "'" & myUpload.form("txtAssDetail" & indx)  & "'"
				assetNo = indx & "-'" & myUpload.form("txtAssDetail" & indx)  & "'"
			else
				totassetNo = totassetNo & ",'" & myUpload.form("txtAssDetail" & indx)  & "'"			
				assetNo = assetNo & "," & indx & "-'" & myUpload.form("txtAssDetail" & indx) 	& "'"
			end if
		next 
		
		assetNo1 = split(assetNo,",")
		
		sqlOra = "select distinct grpcode from fxa_faMst where assetno in (" & totassetNo & ")"
		rsOra.open sqlOra, connOra		

		if not rsOra.EoF then
			do while not rsOra.EoF
				'key = createFlow(connOra,myUpload.form("frmNo"),myUpload.form("orgNo"),myUpload.form("y"),myUpload.form("txtReq"),session("uid"),myUpload.form("txtRemark"))
				sqlOra = "select vOwner from fxa_accassMst where accAsscode='" & rsOra("grpcode") & "'"
				rsOra1.open sqlOra, connOra
				apvno = rsOra1("vOwner")
				rsOra1.Close
				key = createDarFlow(connOra,myUpload.form("frmNo"),myUpload.form("orgNo"),myUpload.form("y"),myUpload.form("txtReq"),session("uid"),myUpload.form("txtRemark"),apvno)

				if key <> false then
					key = split(key, ",")
				else
					response.write "Create request failure!"
					response.end
				end if
				if myUpload.form("selWriteOff") = 2 then
					sqlOra = "delete from flow where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
					& " and cyear2= '" & key(0) & "' and nrunno = '" & key(1) & "' and cstepno in ('23','28')" 					
					rsOra.open sqlOra, connOra		
					sqlOra = "update flow set cstepnextno = '00' where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'" _
					& " and cyear2= '" & key(0) & "' and nrunno = '" & key(1) & "' and cstepno ='08'"
					rsOra.open sqlOra, connOra
				end if
				'for indx = 1 to CInt(myUpload.form("personsMaxline"))
					'sqlOra = "select * from fxa_faMst where assetno = '" & myUpload.form("txtAssDetail" & indx) & "'"
					'rsOra.open sqlOra, connOra
					'if rsOra.EoF then
					'	response.write "Asset No. " & myUpload.form("txtAssDetail" & indx)  & " not found"
					'	response.end
					'end if
					'rsOra.Close
					'response.write "selWriteOff=[" & myUpload.form("selWriteOff" & indx)  & "]<br>"
'					sqlOra = "insert into fxa_darform (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, grpcode, loccode, assetno, nqty, vUnit, writeoffno, vremark) " _
'					& " select '" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "'" _
'					& ", grpcode, loccode, assetno " _
'					& ", '" & myUpload.form("txtQty" & indx) & "', '" & myUpload.form("txtUnit" & indx) & "', '" & myUpload.form("selWriteOff" & indx) & "', '" & myUpload.form("txtRmNo" & indx) & "'" _
'					& " from fxa_faMst where assetno = '" & myUpload.form("txtAssDetail" & indx) & "'"
					
'					& " from fxa_faMst where assetno in (" & assetNo  & ") and grpcode ='" rsOra("grpcode") & "'"
	'					& " from fxa_faMst where assetno = '" & myUpload.form("txtAssDetail" & indx) & "'"
	
'						sqlOra = "insert into fxa_darform (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, grpcode, loccode, assetno, nqty, vUnit, writeoffno, vremark) " _
'						& " select '" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "'" _
'						& ", grpcode, loccode, assetno " _
'						& ", '" & myUpload.form("txtQty" & indx) & "', '" & myUpload.form("txtUnit" & indx) & "', '" & myUpload.form("selWriteOff" & indx) & "', '" & myUpload.form("txtRmNo" & indx) & "'" 

					for indx = 0 to  uBound(assetNo1)
					
						assetNo2 = split(assetNo1(indx),"-")
						'response.write trim(assetNo2(0)) & "<br>"
						'response.write assetNo2(1) & "<br>"
						'& ", '" & myUpload.form("txtQty" & trim(assetNo2(0))) & "', '" & myUpload.form("txtUnit" & trim(assetNo2(0))) & "', '" & myUpload.form("selWriteOff" & trim(assetNo2(0))) & "', '" & myUpload.form("txtRmNo" & trim(assetNo2(0))) & "'"  _
						'sqlOra = "insert into fxa_darform (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, grpcode, loccode, assetno, nqty, vUnit, writeoffno, vremark) " _
						sqlOra = "insert into fxa_darform (nFrmNo, vOrgNo, cYear, cYear2, nRunNo, grpcode, loccode, assetno, nqty, vUnit, writeoffno, vremark) " _
						& " select '" & myUpload.form("frmNo") & "', '" & myUpload.form("orgNo") & "', '" & myUpload.form("y") & "', '" & key(0) & "', '" & key(1) & "'" _
						& ", grpcode, loccode, " & assetNo2(1)  _
						& ", '" & myUpload.form("txtQty" & trim(assetNo2(0))) & "', '" & myUpload.form("txtUnit" & trim(assetNo2(0))) & "', '" & myUpload.form("selWriteOff") & "','" & myUpload.form("txtRmNo" & trim(assetNo2(0))) & "'"  _
						& " from fxa_faMst where assetno = " & assetNo2(1)  & " and grpcode ='" & rsOra("grpcode") & "'"						
						'response.write sqlOra & "<br>"
						rsOra1.open sqlOra, connOra	
						'if  myUpload.form("selWriteOff" & trim(assetNo2(0))) = 2 then
							'sqlOra = "delete from flow where nfrmno = '" & myUpload.form("frmNo") & "' and vorgno = '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "'"
							'& " and cyear2= '" & key(0) & "' and nrunno = '" & key(1) & "' and " 
						'end if
						'sqlOra = "select vOwner from fxa_accassMst where accAsscode='" & rsOra("grpcode") & "'"
						'rsOra1.open sqlOra, connOra
						
						'sqlOra = "update flow set vapvno = '" & rsOra1("vOwner") & "', vrepno = '" & rsOra1("vOwner") & "'" _
						'sqlOra = "update flow set vapvno = '" & apvno & "', vrepno = '" & apvno & "'" _
						'& " where nfrmno= '" & myUpload.form("frmNo") & "' and vorgno= '" & myUpload.form("orgNo") & "' and cyear = '" & myUpload.form("y") & "' and cyear2= '" & key(0) & "' and nrunno ='" & key(1) & "'" _
						'& " and cExtdata = '01'"
'						response.write sqlOra
						'rsOra1.Close
						'rsOra1.open sqlOra, connOra
					next					
					'response.end
					'rsOra1.open sqlOra, connOra	
					'cntLine = cntLine + 1

'''					cntLine = 1
'''					path = server.mapPath("..") & "\dar\file\" & myUpload.form("frmNo") & "_" & myUpload.form("orgNo") & "_" & myUpload.form("y") & "_" & key(0) & "_" &  key(1) & "\"			
'''					createOSDir path
'''					myUpload.path = path
'''					for each multiPart in myUpload.FormEx
'''						if IsObject(myUpload.Form(multiPart)) then
'''							fileName = mid(myUpload.Form(multiPart).UserFilename,instrrev(myUpload.Form(multiPart).UserFilename,"\")+1)
'''							fileName = replace(replace(fileName," ",""),  "'", "-")
'''							myUpload.Form(multiPart).SaveAs path & fileName
			
'''							sqlOra = " insert into fxa_attdarform(nfrmno, vorgno, cyear, cyear2, nrunno, typeno, id, sfile) values('" _
'''							& myUpload.form("frmNo") & "','" & myUpload.form("orgNo") & "','" & myUpload.form("y") & "','" & key(0) & "','" &  key(1) _
'''							& "','" & mid(multiPart,11,1) & "','" & cntLine & "','" & oraFormatStr(fileName) & "')"
'''							'response.write sqlOra & "<br>"
'''							rsOra.open sqlOra, connOra	
'''							cntLine = cntLine + 1
'''						end if
'''					next						
				rsOra.MoveNext
			loop
		end if
	case else
		response.write "Invalid action!"
		response.end
end select
'response.end
connOra.commitTrans
		
'response.end		
if (myUpload.form("bp")) = "" then
	response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/create.asp"
else
	response.redirect (myUpload.form("bp"))
end if


%>
