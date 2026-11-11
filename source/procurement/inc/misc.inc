<%
Dim NEW_LINE
Dim newLine
Dim QUOTE

NEW_LINE = chr(13) & chr(10)
newLine = chr(13) & chr(10)
QUOTE = chr(34)

'####################
'# Function: countChar
'# Arguments:
'#	1. src = Source string
'#	2. char = Char to search
'# Return: Number of "char" found in "src"
'# Description: This function will count "char" in "src" then return number of found.
'####################
function countChar(src, char)
	Dim i, count

	count = 0

	for i = 1 to len(src)
		if mid(src, i, 1) = char then
			count = count + 1
		end if
	next

	countChar = count
end function

function formatCur(val)
	if val = "" or isNull(val) then
		formatCur = null
		exit function
	end if

	'val = "0" & val

	formatCur = right(formatCurrency(val, 2, true, true, true), len(formatCurrency(val, 2, true, true, true)) - 1)
end function

function formatCurColor(val, color, zeroColor)
	if val = "" or isNull(val) then
		formatCurColor = null
		exit function
	end if

	if val = 0 then
		formatCurColor = "<font color=" & zeroColor & ">" & val & "</font>"
	else
		formatCurColor = "<font color=" & color & ">" & val & "</font>"
	end if

end function

function getPath(path, pathSep)
	if path = "" then
		getPath = ""
	else
		if inStrRev(path, pathSep) > 1 then
			getPath = left(path, inStrRev(path, pathSep)-1)
		else
			getPath = path
		end if
	end if
end function
%>