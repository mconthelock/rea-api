<%Option Explicit%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i, selected, page

set rsOra = createObject("adodb.recordset")
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Edit form flow</title></head>
<body>
<form name="frmMain" method="get" action="cgi-bin/editFlowStep.asp">
	<input type="hidden" name="no" value="<%=request.queryString("no")%>">
	<input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
	<input type="hidden" name="y" value="<%=request.queryString("y")%>">
	<input type="hidden" name="step" value="<%=request.queryString("step")%>">
	<input type="hidden" name="stepNext" value="<%=request.queryString("stepNext")%>">
    <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
            <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
            <td bgcolor="#AACCEE" align="center"><b>Edit form flow</b></td>
            <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
        </tr>
        <tr>
            <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
            <td bgcolor="#0000FF">
                <table width="100%" border="0" cellspacing="1" cellpadding="2">
                    <!--########## STEP ##########-->
                    <%
					sqlOra = "select cNo, vName from stepMst where cNo = '" & request.queryString("step") & "'"
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
					%>
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Step:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"><font color="#000000"><%=rsOra("cNo") & ": " & rsOra("vName")%></font></td>
                    </tr>
                    <%
						rsOra.close
					%>
                    <!--########## END STEP ##########-->
                    <!--########## PREVIOUS STEP ##########-->
                    <%
					sqlOra = "select cStart from flowMst where nFrmNo = '" & request.queryString("no") & "' " _
					& "and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' " _
					& "and cStepNo = '" & request.queryString("step") & "'"
					'response.write sqlOra & "<br>"
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
						if rsOra("cStart") = "0" then
							rsOra.close
					
							sqlOra = "select flowMst.*, vName, vDesc " _
							& "from flowMst, stepMst " _
							& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _ 
							& "and cYear = '" & request.queryString("y") & "' and stepMst.cNo = flowMst.cStepNo " _
							& "and stepMst.cNo not in ('--', '00') and stepMst.cNo <> '" & request.queryString("step") & "'" _
							& "order by flowMst.cStepNo"
							'response.write sqlOra & "<br>"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
					%>
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Previous 
                            step:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> 
                            <select name="selPrevStep">
                                <%
								do while not rsOra.EOF
									if rsOra("cStepNextNo") = request.queryString("step") then
										selected = "selected"
									else
										selected = ""
									end if
									
									response.write "<option value=" & rsOra("cStepNo") & " " & selected & ">" & rsOra("cStepNo") & ": " & rsOra("vName") & " - " & rsOra("vDesc") & "</option>"
									rsOra.moveNext
								loop
							else
							end if
							
							rsOra.close
							%>
                            </select>
                        </td>
                    </tr>
                    <%
						else
							rsOra.close
						end if
					else
						rsOra.close
					end if
					%>
                    <!--########## END PREVIOUS STEP ##########-->
                    <!--########## PAGE LOCATION ##########-->
                    <%
					sqlOra = "select vUrl, vFormPage from formMst, flowMst " _
					& "where nFrmNo = '" & request.queryString("no") & "' and nNo = nFrmNo " _
					& "and flowMst.vOrgNo = '" & request.queryString("orgNo") & "' and formMst.vOrgNo = flowMst.vOrgNo " _
					& "and flowMst.cYear = '" & request.queryString("y") & "' and formMst.cYear = flowMst.cYear " _
					& "and cStepNo = '" & request.queryString("step") & "'"
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
						if isNull(rsOra("vUrl")) then
							page = rsOra("vFormPage")
						else
							page = rsOra("vUrl")
						end if
					end if
					%>
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Page 
                            location:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> 
                            <input type="text" name="txtUrl" maxlength="1024" size="50" value="<%=page%>">
                        </td>
                    </tr>
                    <%
					rsOra.close
					%>
                    <!--########## END PAGE LOCATION ##########-->
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
                    <!--########## NEXT STEP ##########-->
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
                    <!--########## APPROVER ##########-->
                    <tr> 
                        <td nowrap width="1%" valign="top" bgcolor="#FFCC00"><b><font color="#000000">Approver:</font></b> 
                        </td>
                        <td nowrap bgcolor="#FFFFCC" valign="top"><font color="#000000"> 
                            <input type="radio" name="optApvType" value="1" <%
							if request.queryString("aType") = "1" then
								response.write "checked"
							end if
							%>>
                            <%
'							sqlOra = "select distinct sPosCode, sPosition from amec.pPosition where sPosCode in (select vPosNo from orgPos where vOrgNo in (select vOrgNo from orgTree start with vorgno = '" & request.queryString("orgNo") & "' connect by vorgno = prior vparent) and vPosNo is not null)"

							sqlOra = "select distinct sPosCode, sPosition from amec.pPosition "
							'response.write sqlOra & "<br>"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
							%>
                            <select name="selPos" onClick="javascript: window.document.frmMain.optApvType[0].checked = true;">
                                <%
							do while not rsOra.EOF
								if request.queryString("aType") = "1" and request.queryString("pos") = rsOra("sPosCode") then
									selected = "selected"
								else
									selected = ""
								end if
								
								response.write "<option value=" & rsOra("sPosCode") & " " & selected & ">" & rsOra("sPosCode") & ": " & rsOra("sPosition") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
                            <%
							end if
							
							rsOra.close
							%>
                            (Refer to requester)<br>
                            </font> 
                            <input type="radio" name="optApvType" value="2" <%
							if request.queryString("aType") = "2" then
								response.write "checked"
							end if
							%>>
                            <%
							sqlOra = "select distinct orgPos.vOrgNo, vName, vPosNo, sPosition from orgPos, amec.pPosition, vOrgMst where vOrgMst.vOrgNo = orgPos.vOrgNo and vPosNo in (select vPosNo from orgPos where vOrgNo in (select vOrgNo from orgTree start with vorgno = '" & request.queryString("orgNo") & "' connect by vorgno = prior vparent)) and orgPos.vOrgNo in (select vOrgNo from orgTree start with vorgno = '50603' connect by vorgno = prior vparent) and sPosCode = vPosNo"

							'18-FEB-03
							'New method is support above case and I am not modify any flow utility, but it very very quite slowly that I got it. :)
							'For more information see addFlowStepFrm.asp at same section.
							sqlOra = "select " _
							& "distinct orgPos.vOrgNo, vName, vPosNo, sPosition " _
							& "from " _
							& "orgPos, amec.pPosition, vOrgMst " _
							& "where " _
							& "vOrgMst.vOrgNo = orgPos.vOrgNo and sPosCode = vPosNo "
							'response.write sqlOra & "<br>"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
							%>
                            <select name="selOwnerPos" onClick="javascript: window.document.frmMain.optApvType[1].checked = true;">
                                <%
							do while not rsOra.EOF
								if request.queryString("aType") = "2" and request.queryString("pos") = rsOra("vPosNo") and request.queryString("orgPos") = rsOra("vOrgNo") then
									selected = "selected"
								else
									selected = ""
								end if
								
								response.write "<option value=" & rsOra("vOrgNo") & "," & rsOra("vPosNo") & " " & selected & ">" & rsOra("vPosNo") & ": " & rsOra("vName") & " " & rsOra("sPosition") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
                            <%
							end if
							
							rsOra.close
							%>
                            <font color="#000000">(Refer to form owner) </font><br>
                            <input type="radio" name="optApvType" value="3" <%
							if request.queryString("aType") = "3" then
								response.write "checked"
							end if
							%>>
                            <input type="text" name="txtApv" maxlength="200" size="10" value="<%=request.queryString("apv")%>" onClick="javascript: window.document.frmMain.optApvType[2].checked = true;">
                            <font color="#000000">(Specific employee ID)</font> 
                        </td>
                    </tr>
                    <!--########## END APPROVER ##########-->
                    <!--########## APPROVAL STYLE ##########-->
                    <tr valign="top"> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Approver</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> 
                            <input type="radio" name="optApvMethod" value="<%=APV_TYPE_SINGLE%>" <%if request.queryString("at") = APV_TYPE_SINGLE then response.write "checked" end if%>>
                            <font color="#000000">Single approver</font><br>
                            <!--
                            <input type="radio" name="optApvMethod" value="<%=APV_TYPE_MULTIPLE_CAN%>" <%if request.queryString("at") = APV_TYPE_MULTIPLE_CAN then response.write "checked" end if%>>
                            <font color="#000000">Multiple approver (Candidate)</font><br>-->
                            <input type="radio" name="optApvMethod" value="<%=APV_TYPE_MULTIPLE_CO%>" <%if request.queryString("at") = APV_TYPE_MULTIPLE_CO then response.write "checked" end if%>>
                            <font color="#000000">Multiple approver</font>&nbsp; 
                            <!--<input type="checkbox" name="chkRejType" value="1" <%if request.queryString("rt") = "1" then response.write "checked" end if%>>
                            <font color="#000000"> Reject flow when someone reject</font>-->
                            <hr width="80%" size="1" align="left">
                            <input type="checkbox" name="chkRejApply" value="<%=APPLY_ALL_REJ%>" <%if request.queryString("a") = APPLY_ALL_REJ or request.queryString("a") = APPLY_ALL_BOTH then response.write "checked" end if%>>
                            <font color="#000000"> Apply approve/reject status 
                            to same approver in other step.</font> </td>
                    </tr>
                    <!--########## END APPROVAL STYLE ##########-->
                    <!--########## REJECT STYLE ##########-->
                    <%if 1 = 2 then%>
                    <tr valign="top"> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Rejector:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> 
                            <input type="radio" name="optRejMethod" value="<%=APV_TYPE_SINGLE%>" <%if request.queryString("rt") = APV_TYPE_SINGLE then response.write "checked" end if%>>
                            <font color="#000000">Single rejector</font><br>
                            <!--
                            <input type="radio" name="optRejMethod" value="<%=APV_TYPE_MULTIPLE_CAN%>" <%if request.queryString("at") = APV_TYPE_MULTIPLE_CAN then response.write "checked" end if%>>
                            <font color="#000000">Multiple approver (Candidate)</font><br>-->
                            <input type="radio" name="optRejMethod" value="<%=APV_TYPE_MULTIPLE_CO%>" <%if request.queryString("rt") = APV_TYPE_MULTIPLE_CO then response.write "checked" end if%>>
                            <font color="#000000">Multiple rejector</font><br>
                        </td>
                    </tr>
                    <%end if%>
                    <!--########## END REJECT STYLE ##########-->
                    <!--########## EXTRA DATA ##########-->
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Extra 
                            data:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"> 
                            <input type="text" name="txtExtData" maxlength="2" size="3" value="<%=request.queryString("extData")%>">
                        </td>
                    </tr>
                    <!--########## END EXTRA DATA ##########-->
                    <%
					else
						rsOra.close
					%>
                    <tr> 
                        <td bgcolor="#FFFFFF" align="center" colspan="2"><br>
                            <font class="alertMsg"><b>Step not found!</b></font><br>
                            &nbsp;</td>
                        <%end if%>
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
                <input type="submit" value="Save change">
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