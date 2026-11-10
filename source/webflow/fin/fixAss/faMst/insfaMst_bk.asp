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
<%
		Response.CharSet = "windows-874" 
		Const ForReading = 1, ForWriting = 2, ForAppending = 3
		Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0	
		
		Dim sqlOra
		Dim rsOra ,myUpload
		Dim  fileName
		Dim Filepath, tmpData
		Dim strNewFileName
		Set rsOra=Createobject("ADODB.Recordset")
		Set myUpload =  server.createObject("SoftArtisans.FileUp")			
		

		
		connOra.BeginTrans
		path = server.mapPath("..") & "\faMst\file\" & year(date) & chkLenth(month(date)) & chkLenth(day(date)) & "\"			
'		response.write path
'		response.end
		createOSDir path
		myUpload.path = path
		
		fileName = mid(myUpload.Form("file").UserFilename,instrrev(myUpload.Form("file").UserFilename,"\")+1)
		fileName = name &"" & fileName
		fileName = replace(replace(fileName," ",""),  "'", "-")
		strNewFileName = Replace(fileName, " ", "_", 1, -1, 1)
		
		fileName = year(date) & chkLenth(month(date)) & chkLenth(day(date)) & chkLenth(Hour(Now)) & chkLenth(Minute(Now))  & "_" & strNewFileName
		
		myUpload.SaveAs path & fileName
		Filepath = path& fileName

		
Set myUpload=Nothing
'response.end
connOra.committrans

'response.redirect(myUpload.Form("bp"))

writeExcelData(Filepath)
response.redirect("faMst.asp")
%>

<%
function writeExcelData(Filepath)

	Dim rs,sql,i

	sql = "SELECT * FROM [Sheet1$];"
	
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

		    Do While Not rs.EOF

		        'Response.Write "<tr>"

		        For I = 0 To rs.Fields.Count - 1
		            'Response.Write "<td>" & rs.Fields.Item(I).Value & "</td>"
					if I = 0 then 
						tmpData = "'" & oraFormatStr(rs.Fields.Item(I).Value) & "'"
					else
						tmpData = tmpData & ",'" & oraFormatStr(rs.Fields.Item(I).Value) & "'"					
					end if
		        Next

		        'Response.Write "</tr>"
				'Response.Write tmpData & "<br>"
				sqlOra = "insert into fxa_tempfaMst values (" & tmpData & ")"
				'Response.Write sqlOra & "<br>"
				rsOra.open sqlOra, connOra
		        rs.MoveNext
		    Loop
		    %>
		    </tbody>
		</table>
		<%

		rs.Close

	end if

	Set rs = Nothing

end function

function runSQL(SQL,rs, Filepath)
	on error resume next

	dim myrs

	set myRs = createobject("ADODB.recordset")
'	myRs.Open SQL,"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("unlocodes.xlsx") & ";Extended Properties=""Excel 8.0;IMEX=1;""", 1, 3
'	myRs.Open SQL,"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("unlocodes.xlsx") & ";Extended Properties=""Excel 8.0;HDR=Yes;""", 1, 3
'	myRs.Open SQL,"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Filepath & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""
	myRs.Open SQL,"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Filepath & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1"""
'sql = "insert into fxa_tempfaMst  select * from  from [Sheet1$]"

'	myRs.Close
'	sql = "insert into fxa_tempfaMst select * from " _
'	& " OPENROWSET('Microsoft.ACE.OLEDB.12.0',Data Source=" & Filepath & ";Extended Properties=""Excel 12.0 Xml;HDR=Yes;IMEX=1; 'SELECT * FROM [ProductList$]');"""
'	myRs.open sqlOra, connOra
    


	set rs = myRs
	if err then
		runSQL = false
		response.write err.description

	else
		runSQL = true
	end if

end function

%>
