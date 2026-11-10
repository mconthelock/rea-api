<% 
	function timeIsOK()
	dim todaysDate
	todaysDate=now()

	if FormatDateTime(todaysDate,4) <= "16:30" then	
		timeIsOK = false
	else
		timeIsOK = true
	end if
end function
%>
