<%
'##########
'# Function: getValue(a, b)
'# Description: Use this function to preventive from repeating displaying data.
'##########
function getValue(a, b)
	if a = b then
		getValue = a
		response.write ""
	else
		getValue = b
		response.write b
	end if
end function

'##########
'# Function: showText(text, num, trail)
'# Description: This function return charactors in text length num charactor and append trail to bottom if text longer than num.
'##########
function showText(text,num, trail)
	if isNull(text) then
		showText = ""
		exit function
	end if

	if text = "" then
		showText = ""
		exit function
	end if

	if len(text) > num then
		showText = left(text, num) & trail
	else
		showText = text
	end if
end function

'#######################
'# Sub: displayPageNavigator
'# Created: 01-MAR-2002
'# Objective: Display page navigator
'# Arguments:
'#	absolutePage = current page
'#	pageCount = cInt(recordCount) devide by pageSize
'#	pageSize = number of record per page
'#	cInt(recordCount) = record count
'#	dest = destination page which will display when click Next and Previous button
'#######################

sub displayPageNavigator(absolutePage, recordCount, pageSize, dest)
	Dim pageCount

	if cSng(recordCount) >= pageSize then
		pageCount = cSng(recordCount)\pageSize
	else
		pageCount = 1
	end if
'		response.write pageSize
'		if cInt(recordCount) > DEFAULT_PAGE_SIZE then
		if cSng(recordCount) > DEFAULT_PAGE_SIZE then
			response.write "<table align=center width=100%" & "><tr valign=top>"
			'Display previous button
			if absolutePage > 1 then
			%>
			<td align="left" width="20%"><input class="inputNormal" type="button" value="Previous" onClick="javascript: window.location.href='<%=dest & absolutePage - 1%>'"></td>
			<%
			else
			%>
			<td align="left" width="20%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<!input type="button" value="Previous"></td>
			<%
			end if
			
			'########## DISPLAY PAGE JUMPER ##########
			if 1 = 1 then
			response.write "<td align=center><font class=" & QUOTE & "FontText" & QUOTE & ">"
			
			if cSng(recordCount) > pageSize then
				for i = 1 to cSng(recordCount) step pageSize
					if i\pageSize + 1 = absolutePage then
				%>
					<b><a class="AText" href="javascript: window.location.href='<%=dest & i\pageSize + 1 %>'"><font color="#FF0000"><%=i%></font></a></b>
				<%
					else
				%>
					<a class="AText" href="javascript: window.location.href='<%=dest & i\pageSize + 1 %>'"><%=i%></a>
				<%
					end if
					
					'Display page seperator
					if cSng(recordCount)  - i - pageSize > 0 then
						response.write ", "
					end if
				next
				
				'response.write "<br><br>"
			end if
			end if
			'########## END DISPLAY PAGE JUMPER ##########
			
			'Display total record
'			response.write "<font color=#FF0000>Total&nbsp;<b>" & cSng(recordCount) & "</b></font>&nbsp;record(s)"
			response.write "</font></td>"
			
			'Display next button
			if absolutePage <= cSng(pageCount) and absolutePage * pageSize < cSng(recordCount) then
			%>
			<td align="right" width="20%"><input class="inputNormal" type="button" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" onClick="javascript: window.location.href='<%=dest & absolutePage + 1%>'"></td>
			<%
			else
			%>
			<td align="right" width="20%">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<!input type="button" value="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Next&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"></td>
			<%
			end if
			response.write "</tr>"
			response.write "<tr><td colspan=3 align=center><font class=" & QUOTE & "FontText" & QUOTE & ">Total&nbsp;<b>" & cSng(recordCount) & "</b>&nbsp;record(s)</font></td></tr></table>"
		else 'Display total record only
		%>

<table width="100%" cellspacing="0" border="0" cellpadding="5">
    <tr>
        <td align="center" height="23"><font class="FontText">Total&nbsp;<%=cSng(recordCount)%>&nbsp;record(s)</font></td>
    </tr>
</table>

		<%
		end if
end sub
%>