<%
	for m = 1 to 12
		select case m
			case 1,3,5,7,8,10,12
				enddate = "31"
			case 2
				enddate = "28"
			case 4,6,9,11
				enddate = "30"
		end select
		for d = 1 to enddate
			tmpdate = d & "/"  & m & "/" & year(now)+1
			
			if (WeekDay(tmpdate) = "1") or (WeekDay(tmpdate) = "7") then
				'response.write tmpdate & "-" & WeekDay(tmpdate) & WeekDayName(WeekDay(tmpdate)) & "<br>"
				response.write "insert into holiday values('" & year(now)+1 & chkLenth(m) & chkLenth(d) & "');" & "<br>"
			end if 
		next
	next
	

Function chkLenth(val)
		select case len(val)
			case 1 
				chkLenth = 0 & val
			case else
				chkLenth = val
		end select 
End Function
%>