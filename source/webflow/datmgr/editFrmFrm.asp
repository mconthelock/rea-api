<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->
<%
Dim rsOra, rsOra2, sqlOra
Dim i, selected

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

%><html>
<head> 
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
<script language="JavaScript" src="../js/window.js"></script>
<title>Edit form</title></head>
<body text="#000000" leftMargin="0" topMargin="0" marginHeight="0" marginawidth="0"> 
<!--#include file="../inc/authen.inc"-->
<table border="0" cellspacing="0" cellpadding="0" align="center" awidth="100%" width="100%">
  <tr> 
    <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" awidth="17" height="18"></td>
    <td background="../imgs/border/book/t.gif" awidth="100%" align="center" width="100%"><img src="../imgs/border/book/t.gif" awidth="17" height="18"></td>
    <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" awidth="17" height="18"></td>
  </tr>
  <tr> 
    <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" awidth="17" height="18"></td>
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="1">
        <tr> 
          <td colspan="2" nowrap bgcolor="#AACCEE" align="center"><b><font color="#000000">Edit 
            form</font></b></td>
        </tr>
        <%
		sqlOra = "select to_char(nNo) as nNo, vOrgNo, cYear, vAName, vName, vDesc, vFormPage, vAuthPage, cStatus, to_char(nLifeTime) as nLifeTime from formMst where nNo =  '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "'"
		'response.write sqlOra
		rsOra.open sqlOra, connOra
					
		if not rsOra.EOF then
	%>
        <!--########## FORM DATA ##########-->
        <form name="frmMain" method="get" action="cgi-bin/editFrm.asp">
          <input type="hidden" name="no" value="<%=rsOra("nNo")%>">
          <input type="hidden" name="orgNo" value="<%=rsOra("vOrgNo")%>">
          <input type="hidden" name="y" value="<%=rsOra("cYear")%>">
          <!--
                    <tr> 
                        <td awidth="1%" nowrap><b><font color="#000000">Organization:</font></b></td>
                        <td awidth="100%"> 
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
          <tr align="center"> 
            <td colspan="2"> 
              <table border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr> 
                  <td bgcolor="#0000FF"> 
                    <table border="0" cellspacing="1" cellpadding="2" width="100%">
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Organization:</font></b></td>
                        <td bgcolor="#FFFFCC"><font color="#000000"><%
						sqlOra = "select * from vOrgMst where vOrgNo = '" & request.queryString("orgNo") & "' order by cType desc, vOrgNo"
						rsOra2.open sqlOra, connOra
						
						if not rsOra2.EOF then
							response.write rsOra2("vOrgNo") & ": " & rsOra2("vName") & " (" & rsOra2("vDesc") & ")"
						else
							response.write "<font color=" & QUOTE & "#000000" & QUOTE & ">Organization not found!</font>"
						end if
						
						rsOra2.close
						%></font></td>
                      </tr>
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Name:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <input type="text" name="txtName" value="<%=rsOra("vName")%>" maxlength="60" size="35" title="Form's name">
                        </td>
                      </tr>
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Alias 
                          Name:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <input type="text" name="txtAName" value="<%=rsOra("vAName")%>" maxlength="20" size="35" title="Alias name">
                        </td>
                      </tr>
                      <tr valign="top"> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Description:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <textarea name="txtDesc" cols="40" title="Form's  description" rows="5"><%=rsOra("vDesc")%></textarea>
                        </td>
                      </tr>
		      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Startup 
                          page:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <input type="text" name="txtFormPage" value="<%=rsOra("vFormPage")%>" maxlength="1024" size="45">
                        </td>
                      </tr>        
		      
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Visibility</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <input type="radio" name="cVisible" value="<%=FORM_STS_SHOW%>" <%if rsOra("cStatus") = FORM_STS_SHOW then response.write "checked" end if%>>
                          <font color="#000000">Show</font>&nbsp; 
                          <input type="radio" name="cVisible" value="<%=FORM_STS_HIDE%>" <%if rsOra("cStatus") = FORM_STS_HIDE then response.write "checked" end if%>>
                          <font color="#000000">Hide</font></td>
                      </tr>
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Life 
                          time:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <input type="text" name="txtLifeTime" size="5" maxlength="3" value="<%=rsOra("nLifeTime")%>">
                          <font color="#000000">day(s) </font></td>
                      </tr>
                      <tr> 
                        <td bgcolor="#FFCC00" awidth="1%" width="1%">&nbsp;</td>
                        <td bgcolor="#FFFFCC">&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <!--
                    <tr> 
                        <td awidth="1%" nowrap><b><font color="#000000">Authority module:</font></b></td>
                        <td awidth="100%">
                            <input type="text" name="txtAuthPage" value="<%=rsOra("vAuthPage")%>" maxlength="1024" size="30">
                        </td>
                    </tr>
					-->
          <%rsOra.close%>
          <tr align="center"> 
            <td colspan="2"> 
              <input type="submit" value="Save change" title="Save change to database">
              <input type="button" value="Authorization..." onClick="javascript: window.location.href = 'formAuthen.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'">
              <!input type="button" value="Configuration...">
              <input type="button" value="Back" onClick="javascript: window.location.href='frmLst.asp?ownerNo=<%=request("ownerNo")%>'">
            </td>
          </tr>
          <tr> 
            <td colspan="2" nowrap height="4"><img src="../imgs/spacer.gif" height="4"></td>
          </tr>
        </form>
        <%else%>
        <tr> 
          <td align="center" bgcolor="#FFFFFF" colspan="2"><br>
            <font class="alertMsg"><b>Form not found!</b></font><br>
            <br>
            <input type="button" value="Back" onClick="javascript: window.location.href = 'frmLst.asp?ownerNo=<%=request("ownerNo")%>'">
          </td>
        </tr>
        <%
					rsOra.close
				end if
				%>

        <!--########## END FORM DATA ##########-->
        <%if session("uid") = "01026" and 1 = 2 then%>
        <!--########## ORG FLOW ##########-->
        <form name="frmOrgFlow">
          <tr> 
            <td  bgcolor="#FFFFFF" colspan="2"><b><font color="#0000FF">Organize 
              Flow</font></b></td>
          </tr>
          <tr> 
            <td awidth="100%" align="center" colspan="2"> 
              <table awidth="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td nowrap valign="top" bgcolor="#0000FF"> 
                    <table awidth="100%" border="0" cellspacing="1" cellpadding="2">
                      <tr> 
                        <td bgcolor="#FFCC00" awidth="1%" nowrap><b><font color="#000000">Organize 
                          :</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <%
										sqlOra = "select * from vOrgMst where vOrgNo <> '00' order by vOrgNo"
										rsOra.open sqlOra, connOra
										
										if rsOra.EOF then
										else
										%>
                          <select name="selOrg" onChange="javascript: window.location.href = 'editFrmFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&org=' + document.frmOrgFlow.selOrg.options[document.frmOrgFlow.selOrg.selectedIndex].value">
                            <%
											do while not rsOra.EOF
										%>
                            <option value="<%=rsOra("vOrgNo")%>" <%if request.queryString("org") = rsOra("vOrgNo") then response.write "selected" end if%>><%=rsOra("vName")%></option>
                            <%
												rsOra.moveNext
											loop
										%>
                          </select>
                          <!input type="button" value="Edit..." onClick="javascript: window.location.href = 'orgStepFrm.asp?no=<%=request.queryString("no")%>
                          &orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&org=' + document.frmOrgFlow.selOrg.options[document.frmOrgFlow.selOrg.selectedIndex].value"> 
                          <%
										end if
										
										rsOra.close
										%>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td awidth="100%"  bgcolor="#FFFFFF" colspan="2"><img src="/sys/imgs/spacer.gif" height="4"></td>
          </tr>
          <tr> 
            <td colspan="2" nowrap valign="top"> 
              <table awidth="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td nowrap valign="top" bgcolor="#0000FF"> 
                    <table awidth="100%" border="0" cellspacing="1" cellpadding="2">
                      <tr class="hdrRow"> 
                        <td nowrap awidth="1%" align="center"><b><font color="#000000">No</font></b></td>
                        <!--<td nowrap awidth="1%" align="center"><b><font color="#000000">Level</font></b></td>-->
                        <td nowrap awidth="1%" align="center"><b><font color="#000000">Step 
                          No</font></b></td>
                        <td nowrap awidth="1%" align="center"><b><font color="#000000">Name</font></b></td>
                        <td align="center"><b><font color="#000000">Desc</font></b></td>
                        <td nowrap awidth="1%" align="center"><font color="#000000"><b>Approver</b></font></td>
                        <td nowrap awidth="1%" align="center"><font color="#000000"></font></td>
                        <!--<td nowrap awidth="1%"><b><font color="#000000"></font></b></td>-->
                        <td nowrap awidth="1%" align="center">&nbsp;</td>
                      </tr>
                      <%
											sqlOra = "select lv, cStepNo, cStepNextNo, vName, vDesc, vOrgNo, vApvOrgNo, vPosNo, vApvNo, cType, cExtData, cApvType, cRejType, cApplyAll from stepMst, (select level as lv, cStepNo, cStepNextNo, vOrgNo, vApvOrgNo, vPosNo, vApvNo, cType, cExtData, cApvType, cRejType, cApplyAll from orgFlowMst start with nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStart = '1' connect by nFrmNo = prior nFrmNo and vOrgNo = prior vOrgNo and cYear = prior cYear and cStepNo = prior cStepNextNo order by level) where cNo = cStepNo"
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
                        <td nowrap awidth="1%" align="center"><font color="#000000"><%=i%></font></td>
                        <!--<td nowrap awidth="1%" align="center"><font color="#000000"><%=rsOra("lv")%></font></td>-->
                        <td nowrap awidth="1%" align="center"><font color="#000000"><%=rsOra("cStepNo")%></font></td>
                        <td nowrap awidth="1%"><font color="#000000"><%=rsOra("vName")%></font></td>
                        <td><font color="#000000"><%=rsOra("vDesc")%></font></td>
                        <td nowrap align="center"> 
                          <%
												select case rsOra("cType")
													case "1"
													case "2"
														sqlOra = "select vEmpNo from orgPos where vOrgNo in (select distinct vOrgNo from orgTree start with vOrgNo = '" & rsOra("vApvOrgNo") & "' connect by vOrgNo = prior vParent) and vPosNo = '" & rsOra("vPosNo") & "'"
														rsOra2.open sqlOra, connOra
														
														if not rsOra2.EOF then
															do while not rsOra2.EOF
																response.write "<a href=" & QUOTE & "javascript: var winUsrInfo = openUsrInfoWin('../usrInfo.asp?uid=" & rsOra2("vEmpNo") & "'); /*winUsrInfo.fucus();*/ void(0)" & QUOTE & ">" & rsOra("vPosNo") & " (" & rsOra2("vEmpNo") & ")</a><br>"
																rsOra2.moveNext
															loop
														end if
														
														rsOra2.close
													case "3"
														response.write "<a href=../usrInfo.asp?uid=" & rsOra("vApvNo") & ">" & rsOra("vApvNo") & "</a>"
												end select
												%>
                        </td>
                        <td align="center" nowrap awidth="1%"> <font color="#000000"> 
                          <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editFlowStepFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>&pos=<%=rsOra("vPosNo")%>&orgPos=<%=rsOra("vApvOrgNo")%>&apv=<%=rsOra("vApvNo")%>&aType=<%=rsOra("cType")%>&extData=<%=rsOra("cExtData")%>&at=<%=rsOra("cApvType")%>&rt=<%=rsOra("cRejType")%>&a=<%=rsOra("cApplyAll")%>'">
                          </font> </td>
                        <!--<td align="center" nowrap awidth="1%"><input type="button" value="Edit..."></td>-->
                        <td align="center" nowrap awidth="1%"> <font color="#000000"> 
                          <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this step?\n\n<%="Step No: " & rsOra("cStepNo") & "\nName: " & replace(replace(trim(rsOra("vName")), "'", "\'"), QUOTE, "\" & QUOTE) & "\nDesc: " & replace(replace(trim(rsOra("vDesc") & ""), "'", "\'"), QUOTE, "\" & QUOTE)%>')) { window.location.href = 'cgi-bin/delFlowStep.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>' }">
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
                          <font class="alertMsg"><b>Flow not found!</b></font><br>
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
              <input type="button" value="Add step..." onClick="javascript: window.location.href = 'addFlowStepFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'">
              <!input type="button" value="Auto generate" onClick="javascript: window.location.href = 'addFrmFlowFrm.asp?no=<%=request.queryString("no")%>
              &orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'"> </td>
          </tr>
        </form>
        <!--########## END ORG FLOW ##########-->
        <%end if%>
        <!--########## FORM FLOW ##########-->
        <form name="frmFormFlow" method="get" aaction="cgi-bin/editStep.asp">
          <tr> 
            <td colspan="2" nowrap><b><font color="#0000FF">Work flow</font></b></td>
          </tr>
          <tr> 
            <td colspan="2" nowrap valign="top"> 
              <table awidth="100%" border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr> 
                  <td nowrap valign="top" bgcolor="#0000FF"> 
                    <table awidth="100%" border="0" cellspacing="1" cellpadding="2" width="100%">
                      <tr class="hdrRow"> 
                        <td nowrap awidth="1%" align="center" width="1"><b><font color="#000000">No</font></b></td>
                        <!--<td nowrap awidth="1%" align="center"><b><font color="#000000">Level</font></b></td>-->
                        <td nowrap awidth="1%" align="center" width="1"><b><font color="#000000">Step 
                          No</font></b></td>
                        <td nowrap awidth="1%" align="center"><b><font color="#000000">Name</font></b></td>
                        <td align="center"><b><font color="#000000">Desc</font></b></td>
                        <td nowrap awidth="1%" align="center" width="1"><font color="#000000"><b>Approver</b></font></td>
                        <td nowrap awidth="1%" align="center" width="1"><font color="#000000">&nbsp;</font></td>
                        <!--<td nowrap awidth="1%"><b><font color="#000000"></font></b></td>-->
                        <td nowrap awidth="1%" align="center" width="1">&nbsp;</td>
                      </tr>
                      <%
											'sqlOra = "select to_char(nFrmNo) as nNo, cYear, vOrgNo, cStepNo, cStepNextNo, vPosNo, vApvOrgNo, a.vName, a.vDesc " _
											'& "from formFlowMst, stepMst a, stepMst b " _
											'& "where a.cNo = formFlowMst.cStepNo and b.cNo = formFlowMst.cStepNextNo"
											
											sqlOra = "select " _
											& "lv, cStepNo, cStepNextNo, vName, vDesc, vOrgNo, vApvOrgNo, vPosNo, vApvNo, cType, cExtData, cApvType, cRejType, cApplyAll " _
											& "from " _
											& "stepMst " _
											& ", (select " _
											& "level as lv, cStepNo, cStepNextNo, vOrgNo, vApvOrgNo, vPosNo, vApvNo, cType, cExtData, cApvType, cRejType, cApplyAll " _
											& "from " _
											& "flowMst " _
											& "start with nFrmNo = '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "' and cStart = '1' connect by nFrmNo = prior nFrmNo and vOrgNo = prior vOrgNo and cYear = prior cYear and cStepNo = prior cStepNextNo order by level) " _
											& "where " _
											& "cNo = cStepNo"
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
                        <td nowrap awidth="1%" align="center" width="1"><font color="#000000"><%=i%></font></td>
                        <!--<td nowrap awidth="1%" align="center"><font color="#000000"><%=rsOra("lv")%></font></td>-->
                        <td nowrap awidth="1%" align="center" width="1"><font color="#000000"><%=rsOra("cStepNo")%></font></td>
                        <td nowrap awidth="1%"><font color="#000000"><%=rsOra("vName")%></font></td>
                        <td><font color="#000000"><%=showData(rsOra("vDesc"))%></font></td>
                        <td nowrap align="center" width="1"> 
                          <%
							select case rsOra("cType")
								case "1"
									response.write "&nbsp;"
								case "2"
									sqlOra = "select vEmpNo from orgPos where vOrgNo in (select distinct vOrgNo from orgTree start with vOrgNo = '" & rsOra("vApvOrgNo") & "' connect by vOrgNo = prior vParent) and vPosNo = '" & rsOra("vPosNo") & "'"
									rsOra2.open sqlOra, connOra
									
									if not rsOra2.EOF then
										do while not rsOra2.EOF
											response.write "<a href=" & QUOTE & "javascript: var winUsrInfo = openUsrInfoWin('../usrInfo.asp?uid=" & rsOra2("vEmpNo") & "'); /*winUsrInfo.fucus();*/ void(0)" & QUOTE & ">" & rsOra("vPosNo") & " (" & rsOra2("vEmpNo") & ")</a><br>"
											rsOra2.moveNext
										loop
									else
										response.write "&nbsp;"
									end if
									
									rsOra2.close
								case "3"
									response.write "<a href=../usrInfo.asp?uid=" & rsOra("vApvNo") & ">" & rsOra("vApvNo") & "</a>"
								case else
									response.write "&nbsp;"
							end select
							%>
                        </td>
                        <td align="center" nowrap awidth="1%" width="1"> <font color="#000000"> 
                          <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editFlowStepFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>&pos=<%=rsOra("vPosNo")%>&orgPos=<%=rsOra("vApvOrgNo")%>&apv=<%=rsOra("vApvNo")%>&aType=<%=rsOra("cType")%>&extData=<%=rsOra("cExtData")%>&at=<%=rsOra("cApvType")%>&rt=<%=rsOra("cRejType")%>&a=<%=rsOra("cApplyAll")%>'">
                          </font> </td>
                        <!--<td align="center" nowrap awidth="1%"><input type="button" value="Edit..."></td>-->
                        <td align="center" nowrap awidth="1%" width="1"> <font color="#000000"> 
                          <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this step?\n\n<%="Step No: " & rsOra("cStepNo") & "\nName: " & replace(replace(trim(rsOra("vName")), "'", "\'"), QUOTE, "\" & QUOTE)%>')) { window.location.href = 'cgi-bin/delFlowStep.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>&step=<%=rsOra("cStepNo")%>&stepNext=<%=rsOra("cStepNextNo")%>' }">
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
                          <font class="alertMsg"><b>Flow not found!</b></font><br>
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
              <input type="button" value="Add step..." onClick="javascript: window.location.href = 'addFlowStepFrm.asp?no=<%=request.queryString("no")%>&orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'">
              <!input type="button" value="Auto generate" onClick="javascript: window.location.href = 'addFrmFlowFrm.asp?no=<%=request.queryString("no")%>
              &orgNo=<%=request.queryString("orgNo")%>&y=<%=request.queryString("y")%>'"> </td>
          </tr>
          <tr> 
            <td colspan="2" nowrap height=""><img src="../imgs/spacer.gif" height="4"></td>
          </tr>
        </form>
        <!--########## END FORM FLOW ##########-->
      </table>
    </td>
    <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" awidth="17" height="18"></td>
  </tr>
  <tr> 
    <td align="right" valign="top"><img src="../imgs/border/book/blc.gif" awidth="17" height="18"></td>
    <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" awidth="17" height="18"></td>
    <td align="left" valign="bottom"><img src="../imgs/border/book/brc.gif" awidth="17" height="18"></td>
  </tr>
</table>
<br>
<!--#include file="../inc/copyRight.inc"-->
<br>
&nbsp; 
</body>
</html>
