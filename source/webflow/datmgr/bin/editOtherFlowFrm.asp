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
	<link href="../css/common.css" rel="stylesheet" type="text/css">
	<link href="../css/THCommon.css" rel="stylesheet" type="text/css">
<title>Edit form flow</title></head>
<body>
<form name="frmMain" method="get" action="cgi-bin/editOtherFlow.asp">
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
                    <!--########## STEP MASTER ##########-->
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Step:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC"></td>
                    </tr>
                    <!--########## END STEP MASTER ##########-->
                    <!--########## PREVIOUS STEP ##########-->
                    <%
					sqlOra = "select orgFlowMst.*, vName, vDesc " _
					& "from orgFlowMst, stepMst " _
					& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _ 
					& "and cYear = '" & request.queryString("y") & "' and stepMst.cNo = orgFlowMst.cStepNo " _
					& "and stepMst.cNo not in ('--', '00') and stepMst.cNo <> '" & request.queryString("step") & "'" _
					& "order by orgFlowMst.cStepNo"
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
								response.write "<option value=" & rsOra("cStepNo") & ">" & rsOra("cStepNo") & ": " & rsOra("vName") & " - " & rsOra("vDesc") & "</option>"
								rsOra.moveNext
							loop
							%>
                            </select>
                        </td>
                    </tr>
                    <%
					end if
					
					rsOra.close
					%>
                    <!--########## END PREVIOUS STEP ##########-->
                    <!--########## NEXT STEP ##########-->
                    <%
					sqlOra = "select orgFlowMst.*, vName, vDesc " _
					& "from orgFlowMst, stepMst " _
					& "where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _ 
					& "and cYear = '" & request.queryString("y") & "' and stepMst.cNo = orgFlowMst.cStepNo " _
					& "and stepMst.cNo <> '00' " _
					& "order by orgFlowMst.cStepNo"
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
                    <!--########## PAGE LOCATION ##########-->
					<%
					sqlOra = "select vUrl, vFormPage from formMst, orgFlowMst " _
					& "where nFrmNo = '" & request.queryString("no") & "' and nNo = nFrmNo " _
					& "and orgFlowMst.vOrgNo = '" & request.queryString("orgNo") & "' and formMst.vOrgNo = orgFlowMst.vOrgNo " _
					& "and orgFlowMst.cYear = '" & request.queryString("y") & "' and formMst.cYear = orgFlowMst.cYear " _
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
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Page location:</font></b></td>
                        <td nowrap bgcolor="#FFFFCC">
                            <input type="text" name="txtUrl" maxlength="1024" size="50" value="<%=page%>">
                        </td>
                    </tr>
					<%
					rsOra.close
					%>
                    <!--########## END PAGE LOCATION ##########-->
                    <!--########## APPROVER ##########-->
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Map 
                            to position:</font></b> </td>
                        <td nowrap bgcolor="#FFFFCC"><font color="#000000"></font> 
                            <font color="#000000"> </font> 
                            <%
							sqlOra = "select distinct sPosCode, sPosition from amec.pPosition where sPosCode in (select vPosNo from orgPos where vOrgNo in (select vOrgNo from orgTree start with vOrgNo = '" & request.queryString("orgNo") & "' connect by vOrgNo = prior vparent))"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
							%>
                            <select name="selPos">
                                <%
							do while not rsOra.EOF
								if request.queryString("pos") = rsOra("sPosCode") then
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
                        </td>
                    </tr>
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