<!-- #include file ="../Conn.asp" -->
<%
Dim sqlOra, rsOra, r, r2, fs, f, i

set rsOra = createObject("adodb.recordset")

sqlOra = request("s")

if request("r") = 0 then
	r = 0
else
	r = request("r")
end if

if request("r2") = 0 then
	r2 = 0
else
	r2 = int(request("r")) + int(request("r2"))
end if

rsOra.open sqlOra, conn

if rsOra.EOF then
	response.write "Data not found!"
else
	set fs = createObject("Scripting.FileSystemObject")
	set f = fs.createTextFile(server.mapPath(".") & "\vendertoexcel.csv", true)

	'########## GO TO START RECORD ##########
	if r <> 0 then
		for i = 1 to r
			rsOra.moveNext

			if rsOra.EOF then
				exit for
			end if
		next
	end if
	'########## END GO TO START RECORD ##########
	
	'########## EXPORT ##########
	'Export header
	f.writeLine("NO, VENDER CODE ,VENDER NAME, TERM CODE, CITY, CONTACT, CURRENCY, ADDRESS, TEL NO, FAX NO")
	
	'Export data
	do while not rsOra.EOF
		'response.write r & "," & r2 & "<br>"
		if int(r) >= int(r2) - 1 and int(r2) <> 0 then
			exit do
		end if

		f.writeLine(r+1 & ",""" & rsOra("sVendCode") & """,""" & rsOra("sEName") & """,""" & rsOra("sTermDesc") & """,""" & rsOra("sCity") & """,""" & rsOra("sContact") & """,""" & rsOra("sStdCur") & """,""" & rsOra("sEAddress") & """,""" & rsOra("sTelNo") & """,""" & rsOra("sFax") & """")
		r = r + 1
		rsOra.moveNext
	loop
	'########## END EXPORT ##########

	f.close
end if

rsOra.close
set rsOra = nothing
set f = nothing
set fs = nothing

response.write "<a href=""vendertoexcel.csv"">Right click and click ""Save Target as..."" to download file<br>(คลิ๊กขวาแล้วเลือก Save Target As... เพื่อดาวน์โหลดไฟล์)</a>"
%>