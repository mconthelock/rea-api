<%'Option Explicit%>

<!--#include file="inc/system.inc"-->
<!--#include file="inc/db.inc"-->
<!--#include file="inc/adovbs.inc"-->
<!--#include file="inc/connUtil.inc"-->

<form action="<%=request.queryString("SCRIPT_NAME")%>">
<input type="text" name="sql" value="<%=request.queryString("sql")%>">
    <input type="submit" name="Submit" value="Submit">
</form>

<%
Dim rsOra, sqlOra
Dim field

if request.queryString("sql") = "" then
	response.end
end if

set rsOra = createObject("adodb.recordset")

sqlOra = request.queryString("sql")
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	response.write "<table border=1>"
	'Print field header
	response.write "<tr>"
	for each field in rsOra.fields
		response.write "<td>" & field.name & "</td>"
	next
	response.write "</tr>"

	'Print data
	do while not rsOra.EOF
	response.write "<tr>"
		for each field in rsOra.fields
			response.write "<td>" & field & "</td>"
		next
		
		rsOra.moveNext
	response.write "</tr>"
	loop
	response.write "</table>"
end if

rsOra.close
set rsOra = nothing
%>