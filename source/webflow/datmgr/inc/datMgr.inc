<%
function addhtml(str)
	response.write str
end function

sub buildTree(node, level, last)
	response.write "<tr><td width=""1%"" bgcolor=""#FFFFFF"" nowrap>"

  if (level <> 0) then
	' Add | and whitespace
	dim j
	
	for j = 0 to level - 2 
		if (acode(j) = 0) then
			addhtml("<IMG src=" & QUOTE & "/" & SYSTEM_FOLDER & "/imgs/vertical.gif" & QUOTE & " ALIGN=" & QUOTE & " absbottom " & QUOTE & ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
		else
			addhtml("<IMG src=" & QUOTE & "/" & SYSTEM_FOLDER & "/imgs/blank.gif" & QUOTE & " ALIGN=" & QUOTE & "absbottom" & QUOTE & ">&nbsp;&nbsp;&nbsp;")
		end if
	next

	' Add the appropriate corner piece.
	  addhtml("<IMG src=" & QUOTE & "/" & SYSTEM_FOLDER & "/imgs/corner.gif" & QUOTE & "ALIGN= " & QUOTE & "absbottom" & QUOTE & ">&nbsp;")
	  addhtml("<IMG src=" & QUOTE & "/" & SYSTEM_FOLDER & "/imgs/user.gif" & QUOTE & "ALIGN= " & QUOTE & "absbottom" & QUOTE & ">&nbsp;")
 else
		response.write "<img src=""" & "/" & SYSTEM_FOLDER & "/imgs/user.gif"">&nbsp;"
  end if

	dim elementID
	Dim elementNum
	ElementID = "TreeViewer_" & elementNum

	addhtml("<SPAN ID=" & QUOTE & elementID & QUOTE & ">" & "<font class=" & QUOTE & "FontText" & QUOTE & ">" & node & "</font></SPAN>")
  
  response.write "</td>"
  response.write "<td bgcolor=""#FFFFFF"" width=""100%""><font color=""#0000FF"">" & rsOra("vName") & "</font></td>"
  response.write "<td width=""1%"" bgcolor=""#FFFFFF""><input type=""button"" value=""Edit..."" onClick=""javascript: window.location.href = 'editOrgFrm.asp?orgNo=" & rsOra("vOrgNo") & "&orgName=" & server.urlEncode(rsOra("vName") & "") & "&orgDesc=" & server.urlEncode(rsOra("vDesc") & "") & "'""></td>"
  response.write "<td width=""1%"" bgcolor=""#FFFFFF""><input type=""button"" value=""Delete"" onClick=""javascript: if(confirm('Do you want to delete tree?\n\n\'" &  rsOra("vName") & ": " & replace(replace(rsOra("vDesc"), "'", "\'"), QUOTE, "\" & QUOTE) & "\'')) {window.location.href = 'cgi-bin/delOrg.asp?orgNo=" &  rsOra("vOrgNo") & "'}""></td>"
  response.write "</tr>"
end sub
%>