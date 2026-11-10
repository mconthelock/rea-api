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
<!--#include file="../../inc/folder.inc"-->


<%
		Response.CharSet = "windows-874" 
		Const ForReading = 1, ForWriting = 2, ForAppending = 3
		Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0	
		
		Dim sqlOra
		Dim rsOra ,myUpload
		Dim  fileName
		Dim Filepath, tmpData
		Dim strNewFileName
		Dim rsfa , vorgno , grpcode
		Dim remark , key , vempno
		Dim rsAss2 , repNo , controller
		Dim radType, FExtension
		Set rsOra=Createobject("ADODB.Recordset")
		Set rsfa=Createobject("ADODB.Recordset")
		Set rsOra2=Createobject("ADODB.Recordset")
		Set myUpload =  server.createObject("SoftArtisans.FileUp")			
				
		connOra.BeginTrans
		path = server.mapPath("..") & "\faMst\file\" & year(date) & chkLenth(month(date)) & chkLenth(day(date)) & "\"			
'		response.write path
'		response.end
		createOSDir path
		myUpload.path = path
		
		fileName = mid(myUpload.Form("file").UserFilename,instrrev(myUpload.Form("file").UserFilename,"\")+1)
		FExtension = MID(filename,instr(filename,".")+1)
		If FExtension <> "xls" and FExtension <> "xlsx" then
%>
						<script language="javascript">
							alert("Not File Excel")
							history.back();
						</script>
<%		
					response.end
		end if 

		fileName = name &"" & fileName
		fileName = replace(replace(fileName," ",""),  "'", "-")
		strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)
		
		fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName
		myUpload.SaveAs path & fileName
		Filepath = path& fileName
		radType = myUpload.Form("radType")
		
Set myUpload=Nothing
'response.end
connOra.committrans
if radType = "1" then
		writeExcelData(Filepath)
		
		response.end
			
		'sqlOra = "select distinct vempno , fxa_tempfaMst.loccode as loc from  fxa_tempfaMst, fxa_locMst , orgpos " _
		'				&"where fxa_tempfaMst.LOCCODE = fxa_locMst.LOCCODE and fxa_locMst.vorgno = orgpos.vorgno " _
		'				&"and fxa_locMst.sposcode = orgpos.vposno" 
		sqlOra = "select distinct vempno , fxa_tempfaMst.loccode as loc , fxa_accassMst.vowner as controller ,  fxa_tempfaMst.GRPCODE as grp from  fxa_tempfaMst, fxa_locMst , orgpos , fxa_accassMst " _
				&"where fxa_tempfaMst.LOCCODE = fxa_locMst.LOCCODE and fxa_locMst.vorgno = orgpos.vorgno and fxa_locMst.sposcode = orgpos.vposno and fxa_tempfaMst.GRPCODE = fxa_accassMst.ACCASSCODE"					
				rsOra.open sqlOra, connOra
				'response.write sqlOra&"<br>"
		'response.end
		Do While Not rsOra.EOF
					'remark=" "
					'vempno = rsOra("vempno")
					'createfaMstFlow(connOra , session("no") , session("orgno") , session("y") , session("uid") , session("uid") , remark , vempno )
					
					if ISNULL(rsOra("controller")) then
					      controller = ""
					else
						  controller = rsOra("controller")
					end if
					
					key = createfaMstFlow(connOra,session("no"),session("orgno"),session("y"),session("uid"), session("uid"),"",rsOra("vempno"),controller)
					key = split(key, ",")
					' reset step admin
					sqlOra = "select vapvno , vrepno from flow where nfrmno='"&session("no")&"' and vorgno = '"&session("orgno")&"' and cyear = '"&session("y")&"' and cyear2 = '"&key(0)&"' and nrunno = '"&key(1)&"' and cextdata = '10'"
					rsOra2.open sqlOra , connOra
					if rsOra2("vapvno") = session("uid") or rsOra2("vrepno") = session("uid")  then
						sqlOra = "update flow set cstepst = '1', vrealapv = '' , capvstno = '0' , dapvdate = '' , capvtime = '' , vremark = '' where nfrmno='"&session("no")&"' and vorgno = '"&session("orgno")&"' and cyear = '"&session("y")&"' and cyear2 = '"&key(0)&"' and nrunno = '"&key(1)&"' and cextdata = '10'"
						rsfa.open sqlOra , connOra
					end if
					rsOra2.close
					'response.write "y2 = "&key(0) & ", nrunno = " & key(1)&"<br>"
					sqlOra = "select ID from fxa_tempfaMst where LOCCODE = '"&rsOra("loc")&"' and GRPCODE = '"&rsOra("grp")&"' and ID not in (select ID from fxa_locemp)"
					'response.write "ASSET = "&sqlOra&"<br>"
					rsOra2.open sqlOra , connOra
					Do While Not rsOra2.EOF
						sqlOra = "insert into fxa_locemp(NFRMNO , VORGNO , CYEAR , CYEAR2 , NRUNNO , ID) " _
								&"values('"&session("no")&"','"&session("orgno")&"','"&session("y")&"','"&key(0)&"','"&key(1)&"','"&rsOra2("ID")&"')"
						rsfa.open sqlOra , connOra
						rsOra2.MoveNext
					Loop
					rsOra2.close
					'if rsOra("controller") <> "" then
					'	repNo = getRep(session("no"), session("orgno"), session("y"), rsOra("controller"))
					'	if isNull(repNo) then
							'repNo = rsOra2("sempno")
					'		repNo = rsOra("controller")
					'	end if
					'	repNo = ""
					'	sqlOra = "update flow set vapvno = '"&rsOra("controller")&"' , vrepno = '"&repNo&"' where NFRMNO = '"&session("no")&"' and " _
					'			&"VORGNO = '"&session("orgno")&"' and cyear = '"&session("y")&"' and cyear2 = '"&key(0)&"' and NRUNNO='"&key(1)&"' and cextdata='04'"
					'	rsfa.open sqlOra , connOra
					'else
					'	sqlOra = "update flow set "
						
						
						
						
					'end if
					'sqlOra = "select grpcode from  fxa_tempfaMst where loccode = '"&rsOra("loc")&"'"
					'rsPIC.open sqlOra, connOra
					'grpcode = rsPIC("grpcode")
					'rsPIC.close
					'sqlOra = "select count(*) as rnum from fxa_accassMst where accasscode = '"&grpcode&"'"
					'response.write "sqlora ="&sqlOra
					'rsPIC.open sqlOra, connOra
					'if NOT rsPIC.EOF then
					'	response.write "count = "&rsPIC("rnum")&"<br>"
					'end if
					'rsPIC.close
					'sqlOra = "select vowner from fxa_accassMst where accasscode = '"&grpcode&"'"
					'response.write "sqlora = "&sqlOra&"<br>"
					'rsPIC.open sqlOra, connOra
					'if NOT rsPIC.EOF then
					'	response.write "group = "&rsPIC("vowner")&"<br>"
					'end if
					'rsPIC.close
					rsOra.MoveNext
				 Loop
				rsOra.close
				Set rsOra = Nothing
else
	UpdateData(Filepath)

end if


response.redirect "/" & SYSTEM_FOLDER & "/" & FIN_SUB_SYSTEM_FOLDER & "/fixAss/faMst/faMst.asp?radType="&radType

'response.write "testtype = "&myUpload.Form("radType")

'response.redirect("faMst.asp")
%>

<%

function UpdateData(Filepath)
	Dim rs,sql,i, id 
	'Dim rsUD
	'Set rsUD=Createobject("ADODB.Recordset")
	
	sql = "SELECT * FROM [FA MASTER$];"
	if runsql(sql,rs, Filepath) then
		Do While Not rs.EOF
			if oraFormatStr(rs.Fields.Item(0).Value) = "" then
			   Exit do
			else
				sql = "select ASSETNO from fxa_faMst where status <> 'R' and ASSETNO = '"&oraFormatStr(rs.Fields.Item(4).Value)&"'"
				rsOra2.open sql, connOra
				if NOT rsOra2.EOF then
					sql = "update fxa_faMst set GRPCODE = '"&oraFormatStr(rs.Fields.Item(0).Value)&"' , LOCCODE = '"&oraFormatStr(rs.Fields.Item(2).Value)&"' , ASSDESC = '"&oraFormatStr(rs.Fields.Item(5).Value)&"' , " _
							& "DOCDATE = '"&oraFormatStr(rs.Fields.Item(6).Value)&"' , INITIALVALUE = '"&rs.Fields.Item(7).Value&"' , " _
							& "ACCUMDP = '"&rs.Fields.Item(8).Value&"' , " _
							& "BVALUE = '"&rs.Fields.Item(9).Value&"' , INVNO = '"&oraFormatStr(rs.Fields.Item(10).Value)&"' , MODELNO = '"&oraFormatStr(rs.Fields.Item(11).Value)&"' , " _
							& "SNNO = '"&oraFormatStr(rs.Fields.Item(12).Value)&"' , PONO = '"&oraFormatStr(rs.Fields.Item(13).Value)&"' , " _
							& "REFASSET = '"&oraFormatStr(rs.Fields.Item(14).Value)&"' , QTY = '"&rs.Fields.Item(15).Value&"' , UNIT = '"&oraFormatStr(rs.Fields.Item(16).Value)&"' , " _
							& "STATUS = '"&oraFormatStr(rs.Fields.Item(17).Value)&"' , " _
							& "SUPNAME = '"&oraFormatStr(rs.Fields.Item(18).Value)&"' , PRNO = '"&oraFormatStr(rs.Fields.Item(19).Value)&"' , " _
							& "BUDGETNO = '"&oraFormatStr(rs.Fields.Item(20).Value)&"' , REQBY='"&oraFormatStr(rs.Fields.Item(21).Value)&"' , " _
							& "PICNAME = '"&oraFormatStr(rs.Fields.Item(22).Value)&"' , REMARK = '"&oraFormatStr(rs.Fields.Item(23).Value)&"' , PICPATH = '"&oraFormatStr(rs.Fields.Item(24).Value)&"' " _
							& "where ASSETNO = '"&rsOra2("ASSETNO")&"'"
					rsOra.open sql, connOra
					'rsOra.close
					'rsUD.open sql, 
					'response.write "update sql = "&sql&"<br/>"
					
				end if
				rsOra2.close
				'if NOT rsUD.EOF then
				'	response.write "if"
				'else
				'	response.write "else"
				'end if
			end if
			rs.MoveNext
		 Loop
		 rs.close
	end if
	Set rs = Nothing
end function



function writeExcelData(Filepath)

	Dim rs,sql,i , id 

	sql = "SELECT * FROM [FA MASTER$];"
	
	if runsql(sql,rs, Filepath) then
	
		%>
		<table border="1">
		    <thead>
		    	<tr>
			        <%
			        'For I = 0 To rs.Fields.Count - 1
			         '   Response.Write "<th>" & rs.Fields.Item(I).Name & "</th>"
			        'Next
			        %>
		        </tr>
		    </thead>
		    <tbody>
		    <%
				sqlOra = "SELECT NVL(max(id), 0) AS ID FROM fxa_tempfaMst"
				'response.write connOra
				rsOra.open sqlOra, connOra
				id = rsOra("ID")
				rsOra.close
		    Do While Not rs.EOF
		        id = id + 1
		        'Response.Write "<tr>"

		       ' For I = 0 To rs.Fields.Count - 1
		            'Response.Write "<td>" & rs.Fields.Item(I).Value & "</td>"
				'	if I = 0 then 
				'		tmpData = "'" & oraFormatStr(rs.Fields.Item(I).Value) & "'"
				'	else
				'		tmpData = tmpData & ",'" & oraFormatStr(rs.Fields.Item(I).Value) & "'"					
				'	end if
		        'Next

		        'Response.Write "</tr>"
				'Response.Write tmpData & "<br>"
			'response.write rs.Fields.Item(23).Value&"<br>"
				if oraFormatStr(rs.Fields.Item(0).Value) <> "" then
					sqlOra = "insert into fxa_tempfaMst(ID,GRPCODE,LOCCODE,ASSETNO,ASSDESC,DOCDATE,INITIALVALUE,ACCUMDP,BVALUE,INVNO,MODELNO,SNNO,PONO,REFASSET," _
							&"QTY,UNIT,STATUS,SUPNAME,PRNO,BUDGETNO,REQBY,REMARK) values ('" &id& "','"&oraFormatStr(rs.Fields.Item(0).Value)&"','" _ 
							& oraFormatStr(rs.Fields.Item(2).Value) & "','"& oraFormatStr(rs.Fields.Item(4).Value)&"','"&oraFormatStr(rs.Fields.Item(5).Value)&"','" _ 
							& oraFormatStr(rs.Fields.Item(6).Value) & "','"& oraFormatStr(rs.Fields.Item(7).Value)&"','"&oraFormatStr(rs.Fields.Item(8).Value)&"','" _ 
							& oraFormatStr(rs.Fields.Item(9).Value) & "','"& oraFormatStr(rs.Fields.Item(10).Value)&"','"&oraFormatStr(rs.Fields.Item(11).Value)&"','" _ 
							& oraFormatStr(rs.Fields.Item(12).Value) & "','"& oraFormatStr(rs.Fields.Item(13).Value)&"','"&oraFormatStr(rs.Fields.Item(14).Value)&"','" _ 
							& oraFormatStr(rs.Fields.Item(15).Value) & "','"& oraFormatStr(rs.Fields.Item(16).Value)&"','"&oraFormatStr(rs.Fields.Item(17).Value)&"','" _ 
							& oraFormatStr(rs.Fields.Item(18).Value) & "','"& oraFormatStr(rs.Fields.Item(19).Value)&"','"&oraFormatStr(rs.Fields.Item(20).Value)&"','" _ 
							& oraFormatStr(rs.Fields.Item(21).Value) & "','"& oraFormatStr(rs.Fields.Item(23).Value)&"')"		
				   Response.Write sqlOra & "<br>"
				   
					rsOra.open sqlOra, connOra
				else
					Exit do
				end if
		        rs.MoveNext
		    Loop
			
		    %>
		    </tbody>
		</table>
		<%

		rs.Close
	end if
	Set rs = Nothing
   ' response.end
end function

function runSQL(SQL,rs, Filepath)
	on error resume next

	dim myrs

	set myRs = createobject("ADODB.recordset")
'	myRs.Open SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("unlocodes.xlsx") & ";Extended Properties=""Excel 8.0;IMEX=1;""", 1, 3
'	myRs.Open SQL,"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("unlocodes.xlsx") & ";Extended Properties=""Excel 8.0;HDR=Yes;""", 1, 3
'	myRs.Open SQL,"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Filepath & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""
	myRs.Open SQL,"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Filepath & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""
'	sql = "insert into fxa_tempfaMst  select * from  from [Sheet1$]"

'	myRs.Close
'	sql = "insert into fxa_tempfaMst select * from " _
'	& " OPENROWSET('Microsoft.ACE.OLEDB.12.0',Data Source=" & Filepath & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1; 'SELECT * FROM [ProductList$]');"""
'	myRs.open sqlOra, connOra
    


	set rs = myRs
	if err then
		runSQL = false
			'response.write err.description
%>
						<script language="javascript">
							alert("Sheet name is FA MASTER only")
							history.back();
						</script>
<%		
			response.end
	else
		runSQL = true
	end if

end function

%>
