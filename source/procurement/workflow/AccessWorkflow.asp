<% 
'		sAuthor5   == > Ability in P/R
'		sAuthor6   == > Ability in Approve (Don't Use)
'		sAuthor8   == > Ability in P/R
'		sAuthor9   == > Ability in Report

'		sAuthor5=mid(Session("sAuthor"),13,3)
'		sAuthor6=mid(Session("sAuthor"),16,3)
'		sAuthor8=mid(Session("sAuthor"),22,3)
'		sAuthor9=mid(Session("sAuthor"),25,3)
		if mid(Session("sAuthor"),13,3) = "NOT" Then
			Response.redirect "../Security/Notaccess.asp"
		end if 

		if mid(Session("sAuthor"),16,3) = "NOT" Then
			Response.redirect "../Security/Notaccess.asp"
		end if 

		if mid(Session("sAuthor"),22,3) = "NOT" Then
			Response.redirect "../Security/Notaccess.asp"
		end if 

		if mid(Session("sAuthor"),25,3) = "NOT" Then
			Response.redirect "../Security/Notaccess.asp"
		end if 
%>

