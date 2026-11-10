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
	<link href="../css/common.css" rel="stylesheet" type="text/css">
	<link href="../css/THCommon.css" rel="stylesheet" type="text/css">
<title>Add form step</title></head>
<body>
<form name="frmMain" method="get" action="cgi-bin/addOtherFlow.asp">
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
						sqlOra = "select * from stepMst where cNo not in ('--', '00') and cNo not in (select cStepNo from orgFlowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' union select cStepNo from orgFlowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "') order by cNo"
						rsOra.open sqlOra, connOra
						%>
                        <td nowrap bgcolor="#FFFFCC"> 
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
						%></td>
                    </tr>
					<!--########## END STEP MASTER ##########-->
                    <!--########## PAGE LOCATION ##########-->
					<%
					sqlOra = "select vFormPage from formMst " _
					& "where nNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' " _
					& "and cYear = '" & request.queryString("y") & "'"
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
						page = rsOra("vFormPage")
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
                        <td nowrap width="1%" bgcolor="#FFCC00"><font color="#000000"><b>Approver:</b></font></td>
                        <td nowrap bgcolor="#FFFFCC"><font color="#000000"> </font> 
                            <%
							sqlOra = "select distinct sPosCode, sPosition from amec.pPosition where sPosCode not in (select vPosNo from orgFlowMst where nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "')"
							rsOra.open sqlOra, connOra
							
							if not rsOra.EOF then
							%>
                            <select name="selPos">
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
                        </td>
                    </tr>
                    <!--########## END APPROVER ##########-->
					
					<!--########## PREVIOUS STEP ##########-->
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
                    <tr> 
                        <td nowrap width="1%" bgcolor="#FFCC00"><b><font color="#000000">Previous step:</font></b></td>
						
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
					<!--########## END PREVIOUS STEP ##########-->
					<%
					end if
					
					rsOra.close
					%>

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
					<!--########## END NEXT STEP ##########-->
					<%
					end if
					
					rsOra.close
					%>
					
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