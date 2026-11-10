<% Option Explicit %>

<!--#include file="inc/adovbs.inc"-->
<!--#include file="inc/db.inc"-->
<!--#include file="inc/connUtil.inc"-->

<%
Dim rsOra, sqlOra

set rsOra = createObject("adodb.recordset")

sqlOra = "select amec.aEmployee.*, sSection, sDepartment, sDivision, sPosition from amec.aemployee, amec.pSection, amec.pDepartment, amec.pDivision, amec.pPosition where sEmpNo = '" & request.queryString("uid") & "' and amec.pSection.sSecCode = amec.aemployee.sSecCode and amec.pDepartment.sDepCode = amec.aemployee.sDepCode and amec.pDivision.sDivCode = amec.aemployee.sDivCode and amec.pPosition.sPosCode = amec.aemployee.sPosCode"
'response.write sqlOra & "<br>"
on error resume next
rsOra.open sqlOra, connOra
on error goto 0

if connOra.errors.count > 0 then
	response.write "(" & connOra.errors(1).number & ") " & connOra.errors(1).description
'	response.end
end if
%>	

<html>
<head>
<title>Approval's remark</title>
<meta http-equiv=" content-Type" content="text/html; charset=">
<link rel="stylesheet" type="text/css" href=" css/common.css">
<link rel="stylesheet" type="text/css" href=" css/THCommon.css">
</head>
<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2" marginwidth="2" marginheight="2">
<%if not rsOra.EOF then%>
<table border="0" width="100%" height="100%">
<tr>
	<td align="center" valign="top">
      <table border="0" cellspacing="0" cellpadding="0" align=" center" width="100%">
        <tr  valign="bottom"> 
        <td align="left" valign="top" bgcolor="#FFD553"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/imgs/border/procurement/y01.jpg" width="12" height="12"></td>
          <td bgcolor="#FFD553" width="100%">&nbsp;</td>
        <td align="right" bgcolor="#FFD553" valign="top"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/imgs/border/procurement/y02.jpg" width="12" height="12"></td>
    </tr>
	
	<!--########## EMPLOYEE ##########-->
	<!--
    <tr > 
        <td bgcolor="#FFD553" align="right">&nbsp;</td>
        <td awidth="96%" align=" center" valign="top"> 
            <table awidth="100%" border="0" cellspacing="0" cellpadding="0">
                <tr abgcolor="#F8D7D1" align=" center" valign="top"> 
                    <td> 
                        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#F8D7D1">
                            <tr>
                                <td rowspan="8" height="2" align=" center" valign="top"> 
                                    <img src="http://webflow/images/emp/<%=rsOra("sEmpNo")%>.jpg" border="1" width="236" height="157"> 
                                </td>
                                <td rowspan="8" height="0" nowrap valign="top" align="left"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/imgs/spacer.gif" width="4"></td>
                                <td width="1%" height="0" nowrap valign="top" align="left"><b><font class="FontTextBlue">Employee ID : </font></b></td>
                                <td width="100%" height="0" valign="top"><font class="FontText"><%=rsOra("sEmpNo")%></font></td>
                            </tr>
                            <tr> 
                                <td width="1%" nowrap valign="top" align="left" height="0"><b><font class="FontTextBlue">Name 
                                    : </font></b></td>
                                <td width="100%" valign="top" height="0"><font class="FontText" color="#006699"><%=rsOra("sName")%></font></td>
                            </tr>
                            <tr> 
                                <td width="1%" nowrap valign="top" align="left" height="0">&nbsp;</td>
                                <td width="100%" valign="top" height="0"><font class="THFontText">(<%=rsOra("sTName")%>)</font></td>
                            </tr>
                            <tr> 
                                <td width="1%" nowrap valign="top" align="left" height="0"><b><font class="FontTextBlue">Division 
                                    : </font></b></td>
                                <td width="100%" valign="top" height="0"> <font class="FontText" cOLOR="#006699"><%="(" & rsOra("sDivCode") & ") " & rsOra("sDivision")%></font></td>
                            </tr>
                            <tr> 
                                <td width="1%" nowrap valign="top" align="left" height="0"><b><font class="FontTextBlue">Department 
                                    : </font></b></td>
                                <td width="100%" valign="top" height="0"><font class="FontText" color="#006699"><%="(" & rsOra("sDepCode") & ") " & rsOra("sDepartment")%></font></td>
                            </tr>
                            <tr> 
                                <td width="1%" nowrap valign="top" align="left" height="0"><b><font class="FontTextBlue">Section 
                                    : </font></b></td>
                                <td width="100%" valign="top" height="0"><font class="FontText" color="#006699"><%="(" & rsOra("sSecCode") & ") " & rsOra("sSection")%></font></td>
                            </tr>
                            <tr> 
                                <td width="1%" nowrap valign="top" align="left" height="0"><b><font class="FontTextBlue">Position 
                                    : </font></b></td>
                                <td width="100%" valign="top" height="0"><font class="FontText" color="#006699"><%=rsOra("sPosition")%></font></td>
                            </tr>
                            <tr> 
                                <td  align="left"><font class="FontTextBlue"><b>E-Mail:</b></font></td>
                                <td ><a class="AText" href="mailto:<%=rsOra("sRecMail")%>">
                                    <%if rsOra("sRecMail") = "" then response.write "- None -" else response.write rsOra("sRecMail") end if%>
                                    </a></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
        <td bgcolor="#FFD553" align="left">&nbsp;</td>
    </tr>
	
	-->
	<%rsOra.close%>
	<!--########## END EMPLOYEE ##########-->
	
	<!--########## REMARK ##########-->
	<%
	sqlOra = "select vRemark from flow where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cYear2 = '" & request.queryString("y2") & "' and nRunNo = '" & request.queryString("runNo") & "' and cStepNo = '" & request.queryString("step") & "'"
	rsOra.open sqlOra, connOra
	
	if not rsOra.EOF then
	%>
	
    <tr  valign="top"> 
        <td align="right" bgcolor="#FFD553">&nbsp;</td>
                    <td bgcolor="#FFD553"><font color="#000000"><b>Remark:</b></font><br>
                        <%=rsOra("vRemark")%></td>
        <td align="left" bgcolor="#FFD553">&nbsp;</td>
    </tr>
	
	<%
	end if
	
	rsOra.close
	%>
	<!--########## END REMARK ##########-->
	
    <tr  valign="top"> 
        <td align="right" bgcolor="#FFD553">&nbsp;</td>
        <td align="right" bgcolor="#FFD553"> 
            <hr size="1"><a class="AText" href="javascript: close(); void(0)">Close</a></td>
        <td align="left" bgcolor="#FFD553">&nbsp;</td>
    </tr>
    <tr  valign="top"> 
        <td align="left" bgcolor="#FFD553" valign="bottom"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/imgs/border/procurement/y03.jpg" width="12" height="12"></td>
        <td bgcolor="#FFD553">&nbsp;</td>
        <td align="right" valign="bottom" bgcolor="#FFD553"><img src="http://<%=request.serverVariables("SERVER_NAME")%>/<%=SYSTEM_FOLDER%>/imgs/border/procurement/y04.jpg" height="12" width="12"></td>
    </tr>
</table>
	</td>
</tr>
</table>	
<!--########## END EMPLOYEE ##########-->
<%else%>

<font class="alertMsg"><b>Employee not found!</b></font>

<%end if%>
	
</body>
</html>


