<% 
'		sAuthor9   == > Ability in Report

'		sAuthor9=mid(Session("sAuthor"),25,3)

		if mid(Session("sAuthor"),25,3) = "NOT" Then
			Response.redirect "../Security/Notaccess.asp"
		end if 
%>

