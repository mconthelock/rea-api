<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/mailUtil.inc"-->
<!--#include file="../../inc/system.inc"-->
<!--#include file="../../inc/connUtil.inc"-->
<!--#include file="../../inc/oraUtil.inc"-->
<!--#include file="../../inc/date.inc"-->
<!--#include file="../../inc/db.inc"-->
<!--#include file="../../inc/formUtil.inc"-->
<!--#include file="../../inc/flowUtil.inc"-->
<!--#include file="../../inc/dirUtil.inc"-->
<!--#include file="../../qa/rq/inc/chkLen.inc"-->
<!--#include file="../../inc/authen.inc"-->
<!--#include file="../../inc/misc.inc"-->
<%
	Response.CharSet = "windows-874" 
	Const ForReading = 1, ForWriting = 2, ForAppending = 3
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0	
	Dim  sqlOra	
	Dim rsOra ,myUpload
	Dim  fileName, name, data, msg
	Dim FSO
	Dim Filepath, path, multiPart, element
	Dim strNewFileName
	Dim frmNo, orgNo, y, tmprunno, runno,i
	Dim file    
	Dim TextStream		
	Dim Line
	
	Set rsOra=Createobject("ADODB.Recordset")
	Set myUpload =  server.createObject("SoftArtisans.FileUp")			
	Set FSO = server.createObject("Scripting.FileSystemObject")

	yy = year(date)
	mm = month(date)
	dd = day(date)

	connOra.BeginTrans

	path = server.mapPath("..") & "\faMst\file\" & yy & mm &  dd & "\"			

	createOSDir path
	myUpload.path = path

'########################################	Begin upload share load part file	############################################
	fileName = mid(myUpload.Form("SHfile").UserFilename,instrrev(myUpload.Form("SHfile").UserFilename,"\")+1)
	fileName = name &"" & fileName
	fileName = replace(replace(fileName," ",""),  "'", "-")
	strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)
	fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName
	myUpload.Form("SHfile").SaveAs path & fileName
	myUpload.SaveAs path& fileName
	Filepath = path& fileName
						
	if FSO.FileExists(Filepath) Then	

		Set file = FSO.GetFile(Filepath)
		Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)
	
		Do While Not TextStream.AtEndOfStream	

			Line = TextStream.readline
			Line = Line & vbCRLF
			data = split(Line,"|")

			sqlOra = "insert into shrdLoadPrt(nfrmno, vorgno, cyear, cyear2, nrunno, id, itemno, dwgno, desript, prodEff, nqty, ncost, ncostot, nts, vreason ) " _
			& " values ('" & frmNo & "','" & orgNo & "','" & y & "','" & runno(0) & "','" &  runno(1) & "','" & oraFormatStr(data(0)) & "'" _
			& ",'" & data(1) & "','" & oraFormatStr(data(2)) & "','" & oraFormatStr(data(3)) & "','" & oraFormatStr(data(4)) & "','" & data(5) & "','" & data(6) & "'" _
			& ",'" & data(7) & "','" & oraFormatStr(trim(data(8))) &"','" & oraFormatStr(trim(data(9))) &"')" 
			rsOra.open sqlOra, connOra		
		Loop
'		respponse.end
		TextStream.Close
		
		Set TextStream = nothing   
		
	Else
		Response.Write "<h3><i><font color=red> File " & Filename & " does not exist</font></i></h3>"
	End If
'########################################	End upload share load part file	############################################	

'########################################	Begin uploadTS Summary file	############################################
	fileName = mid(myUpload.Form("TSfile").UserFilename,instrrev(myUpload.Form("TSfile").UserFilename,"\")+1)
	fileName = name &"" & fileName
	fileName = replace(replace(fileName," ",""),  "'", "-")
	strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)
	
	fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName
	myUpload.Form("TSfile").SaveAs path & fileName

	Filepath = path& fileName
						
	if FSO.FileExists(Filepath) Then	
		Set file = FSO.GetFile(Filepath)
		Set TextStream = file.OpenAsTextStream(ForReading, TristateUseDefault)
	
		Do While Not TextStream.AtEndOfStream	

			Line = TextStream.readline
			Line = Line & vbCRLF
			data = split(Line,"|")

			sqlOra = "insert into shrdLoadSumTS(nfrmno, vorgno, cyear, cyear2, nrunno, id, mcName, totTS ) " _
			& " values ('" & frmNo & "','" & orgNo & "','" & y & "','" & runno(0) & "','" &  runno(1) & "','" & data(0) & "','" & oraFormatStr(data(1)) & "'" _
			& ",'" & Replace(trim(data(2)),vbNewLine,"") & "')" 
			'response.write sqlOra & "<br>"
			rsOra.open sqlOra, connOra		
		Loop
		'response.end
		Set TextStream = nothing   
	Else
		Response.Write "<h3><i><font color=red> File " & Filename & " does not exist</font></i></h3>"
	End If
'########################################	End uploadTS Summary file############################################
	
'############################	Begin insert attached file	########################################################
	for each multiPart in myUpload.Form
		for each element in myUpload.FormEx(multiPart)

			if IsObject( element ) then
				if not ( element.IsEmpty ) then
					fileName = mid(myUpload.Form(multiPart).UserFilename,instrrev(myUpload.Form(multiPart).UserFilename,"\")+1)
					if multiPart  <> "file" then

						fileName = name &"" & fileName
						fileName = replace(replace(fileName," ",""),  "'", "-")
								
						strNewFileName = oraFormatStr(Replace(fileName, " ", "_", 1, -1, 1))
						fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName
								
						element.SaveAs path& filename
						sqlOra = " insert into attSPFile(nfrmno, vorgno, cyear, cyear2, nrunno, itemno, sfile, sempno) select '" _
						& frmNo & "','" & orgNo & "','" & y & "','" & runno(0) & "','" & runno(1) _
						& "',decode(count(itemno),0,1,max(itemno)+1),'" & oraFormatStr(fileName) & "','" & session("uid") & "'" _
						& " from attSPFile where nfrmno = '" & frmNo & "' and vorgno='" & orgNo & "'" _
						& " and cyear='" & y & "' and cyear2 = '" & runno(0) & "' and nrunno = '" & runno(1) & "'"
						'& " and typeno = '" & typeNo & "'"
						'response.write sqlOra & "<br>"
						'response.end
						rsOra.open sqlOra, connOra
						i = i + 1
					end if
				end if 
			end if
		next
	next
'############################	End insert attached file	########################################################
	Set FSO = nothing
'response.end
connOra.CommitTrans
msg = "is created" & vbcrlf 
call sendMailRequester(frmNo, orgNo, y, runno(0) ,  runno(1) , "06241", msg)		
call sendMailRequester(frmNo, orgNo, y, runno(0) ,  runno(1) , "96244", msg)
	
response.redirect myUpload.form("bp")

%>
