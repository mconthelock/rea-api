<%Option Explicit%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i, page

set rsOra = createObject("adodb.recordset")
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Add form step</title></head>
<body>
<form name="frmMain" method="get" action="cgi-bin/addFlowStep.asp">
	<input type="hidden" name="no" value="<%=request.queryString("no")%>">
	<input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
	<input type="hidden" name="y" value="<%=request.queryString("y")%>">
    <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
            <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
            <td bgcolor="#AACCEE" align="center"><b>Add form flow</b></td>
            <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr>
            <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
            <td bgcolor="#0000FF">
                <table width="100%" border="0" cellspacing="1" cellpadding="2">
                    <!--########## STEP MASTER ##########-->
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Step:</font></b></td>
                        <%
						sqlOra = "select * from stepMst where cNo not in ('--', '00') and cNo > '0' and cNo not in (select cStepNo from flowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' union select cStepNo from flowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "') order by cNo"
						rsOra.open sqlOra, connOra
						%>
                        <td nowrap bgcolor="#FFFFCC"> <font color="#000000"> 
                            <%if not rsOra.EOF then%>
                            <select name="selStep">
                                <%
							do while not rsOra.EOF
								response.write "<option value=" & rsOra("cNo") & ">" & rsOra("cNo") & ": " & rsOra("vName") & " - " & rsOra("vDesc") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
                            <%
						else
							response.write "Step master not found!"
						end if
						
						rsOra.close
						%>
                            </font></td>
                    </tr>
                    <!--########## END STEP MASTER ##########-->
                    <!--########## PREVIOUS STEP ##########-->
                    <%
					sqlOra = "select flowMst.*, vName, vDesc " _
					& "from flowMst, stepMst " _
					& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _ 
					& "and cYear = '" & request.queryString("y") & "' and stepMst.cNo = flowMst.cStepNo " _
					& "and stepMst.cNo <> '00' " _
					& "order by flowMst.cStepNo"
					'response.write sqlOra & "<br>"
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
					%>
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Previous 
                            step:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> <font color="#000000"> 
                            <select name="selPrevStep">
                                <%
							do while not rsOra.EOF
								response.write "<option value=" & rsOra("cStepNo") & ">" & rsOra("cStepNo") & ": " & rsOra("vName") & " - " & rsOra("vDesc") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
                            </font></td>
                    </tr>
                    <%
					end if
					
					rsOra.close
					%>
                    <!--########## END PREVIOUS STEP ##########-->
                    <!--########## PAGE LOCATION ##########-->
                    <%
					sqlOra = "select vFormPage from formMst " _
					& "where nNo = '" & request.queryString("no") & "'" _
					& "and vOrgNo = '" & request.queryString("orgNo") & "'" _
					& "and cYear = '" & request.queryString("y") & "'"
					rsOra.open sqlOra, connOra
					
					%>
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Page 
                            location:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> <font color="#000000"> 
                            <input type="text" name="txtUrl" maxlength="1024" size="50" value="<%=rsOra("vFormPage")%>">
                            </font></td>
                    </tr>
                    <%
					rsOra.close
					%>
                    <!--########## END PAGE LOCATION ##########-->
                    <!--########## APPROVER ##########-->
                    <tr> 
                        <td nowrap width="1%" valign="top" bgcolor="#FFCC00"><font color="#000000"><b>Approver:</b> 
                            </font></td>
                        <td nowrap bgcolor="#FFFFCC" valign="top"><font color="#000000"> 
                            <input type="radio" name="optApvType" value="1">
                            <%
							sqlOra = "select distinct sPosCode, sPosition from amec.pPosition " 'where sPosCode not in (select vPosNo from flowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and vPosNo is not null)"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
							%>
                            <select name="selPos" onClick="javascript: window.document.frmMain.optApvType[0].checked = true;">
                                <%
							do while not rsOra.EOF
								response.write "<option value=" & rsOra("sPosCode") & ">" & rsOra("sPosCode") & ": " & rsOra("sPosition") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
                            <%
							end if
							
							rsOra.close
							%>
                            (Refer to requester)<br>
                            </font> <font color="#000000"> 
                            <input type="radio" name="optApvType" value="2" checked>
                            <%
							'18-FEB-03
							'Because I fix this orgNo then show only data from 50603, this is from I do this system for SES only
							' , rather use following code. 
							sqlOra = "select " _
							& "distinct orgPos.vOrgNo, vName, vPosNo, sPosition " _
							& "from " _
							& "orgPos, amec.pPosition, vOrgMst " _
							& "where " _
							& "vOrgMst.vOrgNo = orgPos.vOrgNo " _
							& "and vPosNo in (select vPosNo from orgPos where vOrgNo in " _
							& "(select vOrgNo from orgTree start with vorgno = '" & request.queryString("orgNo") & "' " _
							& "connect by vorgno = prior vparent)) " _
							& "and orgPos.vOrgNo in (select vOrgNo from orgTree start with vorgno = '50603' " _ 
							& "connect by vorgno = prior vparent) and sPosCode = vPosNo"

							'18-FEB-03
							'This method is not support a case, start at A dept send to B dept and send to C div, if A and B is same level in tree
							sqlOra = "select " _
							& "distinct orgPos.vOrgNo, vName, vPosNo, sPosition " _
							& "from " _
							& "orgPos, amec.pPosition, vOrgMst " _
							& "where " _
							& "vOrgMst.vOrgNo = orgPos.vOrgNo " _
							& "and vPosNo in (select vPosNo from orgPos where vOrgNo in " _
							& "(select vOrgNo from orgTree start with vorgno = '" & request.queryString("orgNo") & "' " _
							& "connect by vorgno = prior vparent)) " _
							& "and orgPos.vOrgNo in (select vOrgNo from orgTree start with vorgno = (select vOrgNo from formMst where nNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "') " _
							& "connect by vorgno = prior vparent) and sPosCode = vPosNo"
							
							'18-FEB-03
							'New method is support above case and I am not modify any flow utility, but it very very quite slowly that I got it. :)
							sqlOra = "select " _
							& "distinct orgPos.vOrgNo, vName, vPosNo, sPosition " _
							& "from " _
							& "orgPos, amec.pPosition, vOrgMst " _
							& "where " _
							& "vOrgMst.vOrgNo = orgPos.vOrgNo and sPosCode = vPosNo " _
							& "order by vOrgNo, vPosNo"
							'response.write sqlOra & "<br>"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
							%>
                            <select name="selOwnerPos" onClick="javascript: window.document.frmMain.optApvType[1].checked = true;">
                                <%
							do while not rsOra.EOF
								response.write "<option value=" & rsOra("vOrgNo") & "," & rsOra("vPosNo") & ">" & rsOra("vPosNo") & ": " & rsOra("vName") & " " & rsOra("sPosition") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
                            <%
							end if
							
							rsOra.close
							%>
                            (Refer to position) <br>
                            <input type="radio" name="optApvType" value="3">
                            <input type="text" name="txtApv" maxlength="6" size="10" onClick="javascript: window.document.frmMain.optApvType[2].checked = true;">
                            (Refer to specific employee ID)</font> </td>
                    </tr>
                    <!--########## END APPROVER ##########-->
                    <!--########## APPROVAL STYLE ##########-->
                    <tr valign="top"> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Approval 
                            method:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> <font color="#000000"> 
                            <input type="radio" name="optApvMethod" value="<%=APV_TYPE_SINGLE%>" <%if request.queryString("at") = APV_TYPE_SINGLE then response.write "checked" end if%> checked>
                            Single approver<br>
                            <input type="radio" name="optApvMethod" value="<%=APV_TYPE_MULTIPLE_CO%>" <%if request.queryString("at") = APV_TYPE_MULTIPLE_CO then response.write "checked" end if%>>
                            Multiple approver&nbsp; 
                            <!--<input type="checkbox" name="chkRejType" value="1" checked><font color="#000000">Reject when someone reject</font>-->
                            </font> 
                            <hr width="80%" size="1" align="left">
                            <font color="#000000"> 
                            <input type="checkbox" name="chkApplyAll" value="<%=APPLY_ALL_REJ%>" checked>
                            Apply approval/reject status to same approver in other 
                            step. </font> </td>
                    </tr>
                    <!--########## END APPROVAL STYLE ##########-->
                    <!--########## REJECT STYLE ##########-->
                    <%if 1 = 2 then%>
                    <tr valign="top"> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Reject 
                            method:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> <font color="#000000"> 
                            <input type="radio" name="optRejMethod" value="1" <%if request.queryString("rt") = "1" then response.write "checked" end if%> checked>
                            Single approver<br>
                            <input type="radio" name="optRejMethod" value="2" <%if request.queryString("rt") = "2" then response.write "checked" end if%>>
                            Multiple approver</font> </td>
                    </tr>
                    <%end if%>
                    <!--########## END REJECT STYLE ##########-->
                    <!--########## EXTRA DATA ##########-->
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Extra 
                            data:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> <font color="#000000"> 
                            <input type="text" name="txtExtData" maxlength="2" size="3" value="<%=request.queryString("extData")%>">
                            </font></td>
                    </tr>
                    <!--########## END EXTRA DATA ##########-->
                    <!--########## NEXT STEP ##########-->
                    <%
					sqlOra = "select flowMst.*, vName, vDesc " _
					& "from flowMst, stepMst " _
					& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _ 
					& "and cYear = '" & request.queryString("y") & "' and stepMst.cNo = flowMst.cStepNo " _
					& "and stepMst.cNo <> '00' " _
					& "order by flowMst.cStepNo"
					'response.write sqlOra & "<br>"
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
					%>
                    <!--
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Previous step:</font></b></td>
						
                        <td nowrap bgcolor="#FFFFCC"> 
                            <select name="selNextStep">
							<%
							do while not rsOra.EOF
								response.write "<option value=" & rsOra("cStepNo") & ">" & rsOra("cStepNo") & ": " & rsOra("vName") & " - " & rsOra("vDesc") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
						</td>
                    </tr>
					-->
                    <%
					end if
					
					rsOra.close
					%>
                    <!--########## END NEXT STEP ##########-->
                </table>
            </td>
            <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/bg.gif">&nbsp;</td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif"></td>
            <td align="center" background="../imgs/border/book/bg.gif"> 
                <input type="submit" value="Add">
                <input type="button" value="Back" onClick="javascript: window.location.href = 'editFrmFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'">
            </td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="../imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
</form>
<!--#include file="../inc/copyRight.inc"-->
</body>
</html>

<%
set rsOra = nothing
%>