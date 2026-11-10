<%
Function showLoc(locname)
'		sql="select a from sequenceOrg start with headno=(select sempno from amec.aemployee where sseccode='" & request("sec") & "'" _ 
'		& " and sposcode='30') connOraect by prior empno=headno"
		sql="select LOCCODE , LOCNAME from fxa_locMst"
		rs.Open sql,connOra
		if not rs.EOF then
%>
<select id = "<%=locname%>" name = "<%=locname%>" disabled="true">
	<option value=""></option>
<%	do while not rs.EOF %>	
		<option value="<%=rs(0)%>"><%=rs(0)%>-<%=rs(1)%></option>
<%		rs.Movenext
	Loop
%>	
</select>
<% else %> -
<% 	end if  
End Function	
%>