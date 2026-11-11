<%
function poforprint(poNo,prNo)
	dim rsOra, sqlOra
	set rsOra = createObject("adodb.recordset")
	
	'keep po for printed
	sqlOra ="insert into poForPrint(spono,printdate, cstatus,sprno)" _
	& " values('" & poNo & "',to_date('" & day(date) & "/" & month(date) & "/" & year(date) & "','dd/mm/yyyy'),'N','" & prNo & "')"
'	response.write sqlOra
'	response.end
	rsOra.open sqlOra, conn 
end function
%>