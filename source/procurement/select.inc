<%
'Include file description
'This file contain function(s) and/or subroutine(s) which concern in
'making select tag
'
'Subroutine description
'Purpose: Make select tag
'Variables:
'dataList = option's text list (array)
'selAttr = select tag type (drop down (normal)=1, list box=multiple)
'defaultOption = default selected option

Public Sub makeSelTag(dataList, selAttr, defaultOption)
	Dim i, j

	j = 0
	response.write "<select " & selAttr & ">"
	for i = 0 to uBound(dataList)
		response.write "<option value='" & dataList(i, j) & "'"
		if dataList(i, j) = defaultOption then
			response.write " selected>"
		else
			response.write ">"
		end if
		response.write dataList(i, j+1) & "</option>"
	next
	response.write "</select>"
End Sub
%>