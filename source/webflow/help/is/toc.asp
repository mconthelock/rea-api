<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->

<%
Dim i, rsOra, rsOra2, sqlOra
Dim sign, url

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

validateUser
%>

<html>
<head>
<!meta http-equiv="refresh" content="60;url=index.asp">
<link href="/<%=SYSTEM_FOLDER%>/help/css/common.css" rel="stylesheet" type="text/css">
</head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">
<%
'Read manual
sqlOra = "select " _
& "vName, formHelp.*, to_char(nNo) as nNo2 " _
& "from formMst, formHelp " _
& "where " _
& "formMst.nNo = formHelp.nFrmNo and formMst.vOrgNo = formHelp.vOrgNo and formMst.cYear = formHelp.cYear " _
& "and cStatus = '" & FORM_STS_SHOW & "' and formHelp.vOrgNo like '506%' " _
& "order by vName"
rsOra.open sqlOra, connOra

if not rsOra.EOF then
	response.write "<table cellpadding=1 cellspacing=0 border=0>"

	do while not rsOra.EOF
			if inStr(rsOra("vUrl"), "?") = 0 then
				sign = "?"
			else
				sign = "&"
			end if
			
		'Display form name
		if not isNull(rsOra("vFormPage")) and session("uid") <> "" then
		%><tr><td><img src="../images/orange_bullet.gif"></td><td><aa href="<%=rsOra("vUrl") & sign & "no=" & rsOra("nNo2") & "&orgNo=" & rsOra("vOrgNo") & "&y=" & rsOra("cYear") & "&bp=" & "http://" & request.serverVariables("SERVER_NAME") & "/" & SYSTEM_FOLDER & "/" & SUB_SYSTEM_FOLDER & "/" & request.serverVariables("SCRIPT_NAME")%>"><%=rsOra("vName")%></a></td></tr><%
		else
			response.write "<li>" & rsOra("vName") '& "<br>"
		end if

		rsOra.moveNext
	loop

	response.write "</table>"
end if

rsOra.close
%>
<table width="100%" border="0" cellspacing="2" cellpadding="1">
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/calendar/help/index.asp">Calendar user's manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/comAcs/help/index.asp">Computer Accessory Requisition 
      user's manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/comSup/help/index.asp">Computer Programing/Operation/Maintenance 
      Request user's manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/darwin/help/index.asp">DARWIN User Registration user's 
      manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/eMail/help/out-index.asp">E-Mail out-of MELCO Registration 
      user's manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/eMail/help/index.asp">E-Mail Registration user's manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/fileSvr/help/index.asp">File Server User Registration 
      user's manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/psr/help/index.asp">Parts Supply User Registration user's 
      manual</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/intranet/doc/intranet.doc">Intranet Content Input</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/EProcFrm/doc/eProc.doc">Procurement User Registration</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/swReqFrm/doc/swReq.doc">Software Request</a></td>
  </tr>
  <tr> 
    <td width="1%"><img src="../../imgs/pointer.gif" width="10" height="10"></td>
    <td><a href="/form/is/wwwRegFrm/doc/www.doc">WWW User Registration</a></td>
  </tr>
</table>
</body> </html> 
