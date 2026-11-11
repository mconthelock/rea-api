<%
'##########
'# Function: oraFormatStr(str)
'# Description: This function replace unsupport charactors in string that include in sql command.
'# Argument: str = string which want to format.
'##########
function oraFormatStr(str)
	oraFormatStr = replace(str & "", """", "''")
end function

'##########
'# Function: showOraError(connOra)
'# Description: This function show Oracle's connection error and stop process.
'# Add date: 01-APR-04
'# Argument: connOra = Oracle's connection.
'##########
function showOraError(connOra)
	if connOra is nothing then
		exit function
	end if

	if connOra.errors.count > 0 then
		response.write "Error!<br>Reason: (" & connOra.errors(0).number & ") " & connOra.errors(0).description
		response.end
	end if
end function
%>