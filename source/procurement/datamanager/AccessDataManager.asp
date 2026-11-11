<% 
'		sAuthor2   == > Ability in DataManager

'		sAuthor2=mid(Session("sAuthor"),4,3)

		if mid(Session("sAuthor"),4,3) = "NOT" Then
			Response.redirect "../Security/Notaccess.asp"
		end if 
%>

