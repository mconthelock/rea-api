<%
Option Explicit
%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, rsOra2, sqlOra
Dim i, selected

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="../css/common.css" rel="stylesheet" type="text/css">
<link href="../css/THCommon.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
<script language="JavaScript" src="../js/window.js"></script>
<title>Edit form</title></head>

<body text="#000000">
    <table border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
            <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/bg.gif"> 
                
            <table width="100%" border="0" cellspacing="2" cellpadding="1">
                <tr> 
                    <td colspan="2" nowrap bgcolor="#AACCEE" align="center"><b><font color="#000000">Edit 
                        form</font></b></td>
                </tr>
                <%
					sqlOra = "select to_char(nNo) as nNo, vOrgNo, cYear, vName, vDesc, vFormPage, vAuthPage from formMst where nNo =  '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "'"
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
					%>
                <!--########## FORM DATA ##########-->
                <form name="frmMain" method="get" action="cgi-bin/editFrm.asp">
                    <input type="hidden" name="no" value="<%=rsOra("nNo")%>">
                    <input type="hidden" name="orgNo" value="<%=rsOra("vOrgNo")%>">
                    <input type="hidden" name="y" value="<%=rsOra("cYear")%>">
                    <tr> 
                        <td width="1%" nowrap><b><font color="#000000">Form No:</font></b></td>
                        <td width="100%"><%=rsOra("nNo")%> </td>
                    </tr>
					<!--
                    <tr> 
                        <td width="1%" nowrap><b><font color="#000000">Organization:</font></b></td>
                        <td width="100%"> 
                            <%
							sqlOra = "select * from vOrgMst where vOrgNo <> '00' order by cType desc, vOrgNo"
							rsOra2.open sqlOra, connOra
							
							if not rsOra2.EOF then
							%>
                            <select name="selOrg">
                                <%
								do while not rsOra2.EOF
									if rsOra2("vOrgNo") = request.queryString("orgNo") then
										selected = "selected"
									else
										selected = ""
									end if
									
									response.write "<option value=" & QUOTE & rsOra2("vOrgNo") & QUOTE & " " & selected & ">" & rsOra2("vOrgNo") & ": " & rsOra2("vName") & " (" & rsOra2("vDesc") & ")</option>"
									rsOra2.moveNext
								loop
							%>
                            </select>
                            <%
							else
							%>
                            <font color="#000000">Organization not found!</font> 
                            <%
							end if
							
							rsOra2.close
							%>
                        </td>
                    </tr>
					-->
                    <tr> 
                        <td width="1%" nowrap><b><font color="#000000">Organization:</font></b></td>
                        <td width="100%"><%
							sqlOra = "select * from vOrgMst where vOrgNo = '" & request.queryString("orgNo") & "' order by cType desc, vOrgNo"
							rsOra2.open sqlOra, connOra
							
							if not rsOra2.EOF then
								response.write rsOra2("vOrgNo") & ": " & rsOra2("vName") & " (" & rsOra2("vDesc") & ")"
							else
								response.write "<font color=" & QUOTE & "#000000" & QUOTE & ">Organization not found!</font>"
							end if
							
							rsOra2.close
							%></td>
                    </tr>
                    <tr> 
                        <td width="1%" nowrap><b><font color="#000000">Name:</font></b></td>
                        <td width="100%"> 
                            <input type="text" name="txtName" value="<%=rsOra("vName")%>" maxlength="30" size="25" title="Name of form">
                        </td>
                    </tr>
                    <tr> 
                        <td width="1%" nowrap><b><font color="#000000">Descript:</font></b></td>
                        <td width="100%"> 
                            <input type="text" name="txtDesc" value="<%=rsOra("vDesc")%>" maxlength="256" size="25" title="Form  description">
                        </td>
                    </tr>
                    <tr> 
                        <td width="1%" nowrap><b><font color="#000000">Startup 
                            page:</font></b></td>
                        <td width="100%"> 
                            <input type="text" name="txtFormPage" value="<%=rsOra("vFormPage")%>" maxlength="1024" size="30">
                        </td>
                    </tr>
                    <!--
                    <tr> 
                        <td width="1%" nowrap><b><font color="#000000">Authority module:</font></b></td>
                        <td width="100%">
                            <input type="text" name="txtAuthPage" value="<%=rsOra("vAuthPage")%>" maxlength="1024" size="30">
                        </td>
                    </tr>
					-->
                    <%rsOra.close%>
                    <tr align="center"> 
                        <td colspan="2"> 
                            <input type="submit" value="Save change" title="Save change to database">
                            <input type="button" value="Back" onClick="javascript: window.location.href='frmLst.asp'">
                        </td>
                    </tr>
                    <tr> 
                        <td colspan="2" nowrap height="4"><img src="../imgs/spacer.gif" height="4"></td>
                    </tr>
                </form>
                <!--########## END FORM DATA ##########-->
                <!--########## OTHER FLOW ##########-->
                <tr> 
                    <td colspan="2" nowrap valign="top"><b><font color="#000000">Requester 
                        flow :</font></b></td>
                </tr>
                <form name="frmOtherFlow" method="get" aaction="cgi-bin/editStep.asp">
                    <tr> 
                        <td colspan="2" nowrap valign="top"> 
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                    <td nowrap valign="top" bgcolor="#0000FF"> 
                                        <table width="100%" border="0" cellspacing="1" cellpadding="2">
                                            <tr class="hdrRow"> 
                                                <td nowrap width="1%" align="center"><b><font color="#000000">No</font></b></td>
                                                <!--<td nowrap width="1%" align="center"><b><font color="#000000">Level</font></b></td>-->
                                                <td nowrap width="1%" align="center"><b><font color="#000000">Step 
                                                    No</font></b></td>
                                                <td nowrap width="1%" align="center"><b><font color="#000000">Name</font></b></td>
                                                <td align="center" width="100%"><b><font color="#000000">Desc</font></b></td>
                                                <td nowrap width="1%" align="center"><font color="#000000"><b>Position</b></font></td>
                                                <td nowrap width="1%" align="center"><font color="#000000"></font></td>
                                                <!--<td nowrap width="1%"><b><font color="#000000"></font></b></td>-->
                                                <td nowrap width="1%" align="center">&nbsp;</td>
                                            </tr>
                                            <%
											sqlOra = "select lv, cStepNo, cStepNextNo, vName, vDesc, vOrgNo, vPosNo, sPosition from stepMst, amec.pPosition, (select level as lv, cStepNo, cStepNextNo, vOrgNo, vPosNo from orgFlowMst start with nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "'  and cStart = '1' connect by nFrmNo = prior nFrmNo and vOrgNo = prior vOrgNo and cYear = prior cYear and cStepNo = prior cStepNextNo order by level) where cNo = cStepNo and sPosCode = vPosNo"
											'response.write sqlOra & "<br>"
											rsOra.open sqlOra, connOra
			
											if not rsOra.EOF then
												i = 1
												'if not rsOra.EOF then
												do while not rsOra.EOF
											%>
                                            <tr bgcolor="<%
									if i mod 2 = 0 then
										response.write "#FFCC00"
									else
										response.write "#FFFFCC"
									end if
									%>"> 
                                                <td nowrap width="1%" align="center"><font color="#000000"><%=i%></font></td>
                                                <!--<td nowrap width="1%" align="center"><font color="#000000"><%=rsOra("lv")%></font></td>-->
                                                <td nowrap width="1%" align="center"><font color="#000000"><%=rsOra("cStepNo")%></font></td>
                                                <td nowrap width="1%"><font color="#000000"><%=rsOra("vName")%></font></td>
                                                <td><font color="#000000"><%=rsOra("vDesc")%></font></td>
                                                <td nowrap><font color="#000000"><%=rsOra("vPosNo") & " (" & rsOra("sPosition") & ")"%></font></td>
                                                <td align="center" nowrap width="1%"><font color="#000000"> 
                                                    <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editOtherFlowFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>&pos=<%=rsOra("vPosNo")%>'">
                                                    </font></td>
                                                <!--<td align="center" nowrap width="1%"><input type="button" value="Edit..."></td>-->
                                                <td align="center" nowrap width="1%"><font color="#000000"> 
                                                    <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this step?\n\n<%="Step No: " & rsOra("cStepNo") & "\nName: " & replace(replace(trim(rsOra("vName")), "'", "\'"), QUOTE, "\" & QUOTE) & "\nDesc: " & replace(replace(trim(rsOra("vDesc")), "'", "\'"), QUOTE, "\" & QUOTE)%>')) { window.location.href = 'cgi-bin/delOtherFlow.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>' }">
                                                    </font></td>
                                            </tr>
                                            <%
												i = i + 1
												rsOra.moveNext
											loop
										else
									%>
                                            <tr> 
                                                <td nowrap align="center" colspan="9" bgcolor="#FFFFFF"><font color="#000000"><br>
                                                    <font class="alertMsg"><b>Flow 
                                                    not found!</b></font><br>
                                                    &nbsp;</font></td>
                                            </tr>
                                            <%
										'end if
									end if
									
									rsOra.close
									%>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr align="center"> 
                        <td colspan="2" nowrap> 
                            <input type="button" value="Add step..." onClick="javascript: window.location.href = 'addOtherFlowFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'">
                            <!input type="button" value="Auto generate" onClick="javascript: window.location.href = 'addFrmFlowFrm.asp?no=<%=request.queryString("no")%>
                            &orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'"> 
                        </td>
                    </tr>
                    <tr> 
                        <td colspan="2" nowrap height=""><img src="../imgs/spacer.gif" height="4"></td>
                    </tr>
                </form>
                <%else%>
                <tr> 
                    <td width="100%" align="center" bgcolor="#FFFFFF"><br>
                        <font class="alertMsg"><b>Form not found!</b></font><br>
                        <br>
                        <input type="button" value="Back" onClick="javascript: window.location.href = 'frmLst.asp'">
                    </td>
                </tr>
                <%
					rsOra.close
				end if
				%>
                <!--########## END OTHER FLOW ##########-->
                <!--########## FORM FLOW ##########-->
                <form name="frmFormFlow" method="get" aaction="cgi-bin/editStep.asp">
                    <tr> 
                        <td colspan="2" nowrap><b><font color="#000000">Form owner 
                            flow :</font></b></td>
                    </tr>
                    <tr> 
                        <td colspan="2" nowrap valign="top"> 
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                <tr> 
                                    <td nowrap valign="top" bgcolor="#0000FF"> 
                                        <table width="100%" border="0" cellspacing="1" cellpadding="2">
                                            <tr class="hdrRow"> 
                                                <td nowrap width="1%" align="center"><b><font color="#000000">No</font></b></td>
                                                <!--<td nowrap width="1%" align="center"><b><font color="#000000">Level</font></b></td>-->
                                                <td nowrap width="1%" align="center"><b><font color="#000000">Step 
                                                    No</font></b></td>
                                                <td nowrap width="1%" align="center"><b><font color="#000000">Name</font></b></td>
                                                <td align="center"><b><font color="#000000">Desc</font></b></td>
                                                <td nowrap width="1%" align="center"><font color="#000000"><b>Approver</b></font></td>
                                                <td nowrap width="1%" align="center"><font color="#000000"></font></td>
                                                <!--<td nowrap width="1%"><b><font color="#000000"></font></b></td>-->
                                                <td nowrap width="1%" align="center">&nbsp;</td>
                                            </tr>
                                            <%
											'sqlOra = "select to_char(nFrmNo) as nNo, cYear, vOrgNo, cStepNo, cStepNextNo, vPosNo, vApvOrgNo, a.vName, a.vDesc " _
											'& "from formFlowMst, stepMst a, stepMst b " _
											'& "where a.cNo = formFlowMst.cStepNo and b.cNo = formFlowMst.cStepNextNo"
											
											sqlOra = "select lv, cStepNo, cStepNextNo, vName, vDesc, vOrgNo, vPosNo, vApvNo from stepMst, (select level as lv, cStepNo, cStepNextNo, vOrgNo, vPosNo, vApvNo from formFlowMst start with nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStart = '1' connect by nFrmNo = prior nFrmNo and vOrgNo = prior vOrgNo and cYear = prior cYear and cStepNo = prior cStepNextNo order by level) where cNo = cStepNo"
											'response.write sqlOra & "<br>"
											rsOra.open sqlOra, connOra
			
											if not rsOra.EOF then
												i = 1
												'if not rsOra.EOF then
												do while not rsOra.EOF
											%>
                                            <tr bgcolor="<%
									if i mod 2 = 0 then
										response.write "#FFCC00"
									else
										response.write "#FFFFCC"
									end if
									%>" valign="top"> 
                                                <td nowrap width="1%" align="center"><font color="#000000"><%=i%></font></td>
                                                <!--<td nowrap width="1%" align="center"><font color="#000000"><%=rsOra("lv")%></font></td>-->
                                                <td nowrap width="1%" align="center"><font color="#000000"><%=rsOra("cStepNo")%></font></td>
                                                <td nowrap width="1%"><font color="#000000"><%=rsOra("vName")%></font></td>
                                                <td><font color="#000000"><%=rsOra("vDesc")%></font></td>
                                                <td nowrap align="center"> 
                                                    <%
												if not isNull(rsOra("vPosNo")) then
													sqlOra = "select vEmpNo from orgPos where vOrgNo in (select distinct vOrgNo from orgTree start with vOrgNo = '" & rsOra("vOrgNo") & "' connect by vOrgNo = prior vParent) and vPosNo = '" & rsOra("vPosNo") & "'"
													rsOra2.open sqlOra, connOra
													
													if not rsOra2.EOF then
														do while not rsOra2.EOF
															response.write "<a href=" & QUOTE & "javascript: var winUsrInfo = openUsrInfoWin('../usrInfo.asp?uid=" & rsOra2("vEmpNo") & "'); /*winUsrInfo.fucus();*/ void(0)" & QUOTE & ">" & rsOra("vPosNo") & " (" & rsOra2("vEmpNo") & ")</a><br>"
															rsOra2.moveNext
														loop
													end if
													
													rsOra2.close
												else
													response.write "<a href=../usrInfo.asp?uid=" & rsOra("vApvNo") & ">" & rsOra("vApvNo") & "</a>"
												end if
												%>
                                                </td>
                                                <td align="center" nowrap width="1%"> 
                                                    <font color="#000000"> 
                                                    <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editFrmFlowFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>&pos=<%=rsOra("vPosNo")%>&apv=<%=rsOra("vApvNo")%>'">
                                                    </font></td>
                                                <!--<td align="center" nowrap width="1%"><input type="button" value="Edit..."></td>-->
                                                <td align="center" nowrap width="1%"> 
                                                    <font color="#000000"> 
                                                    <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this step?\n\n<%="Step No: " & rsOra("cStepNo") & "\nName: " & replace(replace(trim(rsOra("vName")), "'", "\'"), QUOTE, "\" & QUOTE) & "\nDesc: " & replace(replace(trim(rsOra("vDesc")), "'", "\'"), QUOTE, "\" & QUOTE)%>')) { window.location.href = 'cgi-bin/delFrmFlow.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>' }">
                                                    </font></td>
                                            </tr>
                                            <%
												i = i + 1
												rsOra.moveNext
											loop
										else
									%>
                                            <tr> 
                                                <td nowrap align="center" colspan="9" bgcolor="#FFFFFF"><font color="#000000"><br>
                                                    <font class="alertMsg"><b>Flow 
                                                    not found!</b></font><br>
                                                    &nbsp;</font></td>
                                            </tr>
                                            <%
										'end if
									end if
									
									rsOra.close
									%>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr align="center"> 
                        <td colspan="2" nowrap> 
                            <input type="button" value="Add step..." onClick="javascript: window.location.href = 'addFrmFlowFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'">
                            <!input type="button" value="Auto generate" onClick="javascript: window.location.href = 'addFrmFlowFrm.asp?no=<%=request.queryString("no")%>
                            &orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'"> 
                        </td>
                    </tr>
                    <tr> 
                        <td colspan="2" nowrap height=""><img src="../imgs/spacer.gif" height="4"></td>
                    </tr>
                </form>
                <!--########## END FORM FLOW ##########-->
            </table>
            </td>
            <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="../imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
	<br>
<!--#include file="../inc/copyRight.inc"-->
</body>
</html>
