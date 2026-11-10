<%
function showAttFile(no, orgNo, y, y2, runNo, folName, tyno, mode,bp)
	Dim rsOra, sqlOra
	Dim  i, attFile, path
	i = 0
	set rsOra = createObject("adodb.recordset")
	sqlOra = "select sfile from fxa_attdarform where nfrmno = '" & no & "' and vorgno = '" & orgNo & "'" _
	& " and cyear = '" & y & "' and cyear2='" & y2 & "' and nrunno = '" & runNo & "' and typeno = '" & tyno & "' order by id " 
	rsOra.open sqlOra, connOra

	do while not rsOra.EoF
		if i = 0 then
		else
			response.write ", "
		end if
%>
	<a href="http://<%=request.serverVariables("SERVER_NAME") %>/<%=SYSTEM_FOLDER%>/<%=FIN_SUB_SYSTEM_FOLDER%>/<%=folName%>/file/<%=no & "_"  & orgNo & "_" & y & "_" & y2 & "_" &  runNo & "/" & rsOra("sfile")%>" target="_blank"><u><%=rsOra("sfile")%></u></a>
	<%	if (mode=MODE_EDIT)	then	%>	
	<!--(<a href="#" onClick="javascript: if(confirm('Do you want to delete this file?\n\n\File:<%=rsOra("sfile")%>')) {window.location.href = 'delfile.asp?frmNo=<%=no%>&orgNo=<%=(orgNo)%>&y=<%=y%>&y2=<%=y2%>&runNo=<%=runNo%>&sfile=<%=rsOra("sfile")%>&bp=<%=bp%>'}">delete</a>)-->	
<%
			end if	
		rsOra.MoveNext
		i = i+1
	loop
	response.write attFile
	rsOra.Close
end function


%>