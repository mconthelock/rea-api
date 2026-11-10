<%
Const JB_STEP = "01" 'Job Controller

'##########
'# Function: isJobController(no, orgNo, y, y2, runNo, uid)
'# Description: This function check uid in specific transaction for Job Controller position.
'# Argument:
'# 1. Form key.
'# 2. uid = employee id.
'# Return: 
'# 1. True if uid is a Job Controller.
'# 2. False if uid is not a Job Controller.
'##########
function isJobController(no, orgNo, y, y2, runNo, uid)
	Dim sqlOra, rsOra
	
	set rsOra = createObject("adodb.recordset")
	
	sqlOra = "select vApvNo from flow " _
	& "where " _
	& "nFrmNo = '" & no & "' and vOrgNo = '" & orgNo & "' " _
	& "and cYear = '" & y & "' and cYear2 = '" & y2 & "' " _
	& "and nRunNo = '" & runNo & "' " _
	& "and (vApvNo = '" & uid & "' or vRepNo = '" & uid & "') and cExtData = '01' "
'	response.write sqlOra & "<br>"
'	response.end
	rsOra.open sqlOra, connOra
	
	isJobController = not rsOra.EOF
	
	rsOra.close
	set rsOra = nothing
end function

%>