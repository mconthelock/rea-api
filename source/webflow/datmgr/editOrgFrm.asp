<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/misc.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->
<%
Dim rsOra, sqlOra
Dim i, orgType

set rsOra = createObject("adodb.recordset")
%><html>
<head> 
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
<title>Edit organization</title></head>
<body text="#000000" leftMargin="0" topMargin="0" marginHeight="0" marginawidth="0"> 
<%=showPageLoading%> 
<form name="frmMain" method="get" action="cgi-bin/editStep.asp">
  <input type="hidden" name="stepNo" value="<%=request.queryString("stepNo")%>">
  <table border="0" cellspacing="0" cellpadding="0" align="center" awidth="100%" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" awidth="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" awidth="100%" align="center" width="100%"><img src="../imgs/border/book/t.gif" awidth="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" awidth="17" height="18"></td>
    </tr>
    <tr>
      <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Edit tree</font></b></td>
      <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" awidth="17" height="18"></td>
      <td bgcolor="#FFFFFF"> 
        <!--########## ORGANIZATION DATA ###########-->
        <%
		sqlOra = "select * from vOrgMst where vOrgNo = '" & request.queryString("orgNo") & "'"
		rsOra.open sqlOra, connOra
		
		if not rsOra.EOF then
			orgType = rsOra("cType")
		%>
        <table awidth="100%" border="0" cellspacing="2" cellpadding="1" width="100%">
          <tr> 
            <td awidth="1%" nowrap><b><font color="#000000">Org No:</font></b></td>
            <td><font color="#000000"><%=request.queryString("orgNo")%></font></td>
          </tr>
          <tr> 
            <td awidth="1%" nowrap><b><font color="#000000">Name:</font></b></td>
            <td><font color="#000000"><%=rsOra("vName")%></font></td>
          </tr>
          <tr> 
            <td awidth="1%" nowrap><b><font color="#000000">Descript:</font></b></td>
            <td><font color="#000000"><%=rsOra("vDesc")%></font></td>
          </tr>
          <%
		  else
		  	response.write "Organization not found!"
		  end if
		  
		  rsOra.close
		  %>
          <!--########## END ORGANIZATION DATA ###########-->
          <!--########## EMPLOYEE TABLE ##########-->
          <tr> 
            <td colspan="2" nowrap height="4"><img src="../imgs/spacer.gif" height="4"></td>
          </tr>
          <tr> 
            <td colspan="2" nowrap><b><font color="#000000">Employee:</font></b></td>
          </tr>
          <tr> 
            <td colspan="2" nowrap valign="top"> 
              <table awidth="100%" border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr> 
                  <td nowrap valign="top" bgcolor="#0000FF"> 
                    <table awidth="100%" border="0" cellspacing="1" cellpadding="2" width="100%">
                      <tr class="hdrRow" align="center"> 
                        <td nowrap awidth="1%"><b><font color="#000000">No</font></b></td>
                        <td nowrap awidth="1%"><b><font color="#000000">Emp No</font></b></td>
                        <td nowrap awidth="1%"><b><font color="#000000">Name</font></b></td>
                        <td><b><font color="#000000">Position</font></b></td>
                        <!--<td nowrap awidth="1%"><b><font color="#000000"></font></b></td>-->
                        <td nowrap awidth="1%"><b><font color="#000000">&nbsp;</font></b></td>
                      </tr>
                      <%
									sqlOra = "select " _
									& "distinct sEmpNo, sName, amec.aEmployee.sPosCode, orgPos.vPosNo, sPosition " _
									& "from " _
									& "amec.aEmployee, amec.pPosition, orgPos " _
									& "where " _
									& "vOrgNo = '" & request.queryString("orgNo") & "' "  _
									& "and sEmpNo = vEmpNo and amec.pPosition.sPosCode = vPosNo " _
									& "order by sPosition, sEmpNo"
									
									'& "sEmpNo in (select distinct vEmpNo from orgPos where vOrgNo = '" & request.queryString("orgNo") & "') " _
									'& " and  vPosNo in (select distinct vPosNo from orgPos where vOrgNo = '" & request.queryString("orgNo") 									& "') " _
									'& "and orgPos.vEmpNo = sEmpNo and orgPos.vPosNo = amec.pPosition.sPosCode(+) " _
									'& "order by vPosNo, sEmpNo"
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
                        <td nowrap awidth="1%" align="center"><font color="#000000"><%=i%></font></td>
                        <td nowrap awidth="1%" align="center"><font color="#000000"> 
                          <%
						if rsOra("sPosCode") <> rsOra("vPosNo") then
							response.write "(" & rsOra("sEmpNo") & ")"
						else
							response.write rsOra("sEmpNo")
						end if
						%>
                          </font></td>
                        <td nowrap awidth="1%"><font color="#000000"><%=rsOra("sName")%></font></td>
                        <td><font color="#000000"><%=rsOra("sPosition")%></font></td>
                        <!--<td align="center" nowrap awidth="1%"><input type="button" value="Edit..."></td>-->
                        <td align="center" nowrap awidth="1%"> 
                          <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this employee?\n\n<%="Emp No: " & rsOra("sEmpNo") & "\nName: " & replace(replace(trim(rsOra("sName")) & "", "'", "\'") & "", QUOTE, "\" & QUOTE) & "\nPosition: " & replace(replace(trim(rsOra("sPosition")) & "", "'", "\'") & "", QUOTE, "\" & QUOTE)%>')) { window.location.href = 'cgi-bin/delOrgEmp.asp?empNo=<%=rsOra("sEmpNo")%>&orgNo=<%=request.queryString("orgNo")%>&posNo=<%=rsOra("vPosNo")%>' }">
                        </td>
                      </tr>
                      <%
												i = i + 1
												rsOra.moveNext
											loop
										else
									%>
                      <tr> 
                        <td nowrap align="center" colspan="5" bgcolor="#FFFFFF"><br>
                          <font class="alertMsg"><b>Employee not found!</b></font><br>
                          &nbsp;</td>
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
            <td colspan="2" nowrap height=""> 
              <input type="button" value="Add employee..." onClick="javascript: window.location.href = 'addOrgEmpFrm.asp?orgNo=<%=request.queryString("orgNo")%>&orgName=<%=server.urlEncode(request.queryString("orgName") & "")%>&orgDesc=<%=server.urlEncode(request.queryString("orgDesc") & "")%>'">
              <input type="button" value="Automatic addition" onClick="javascript: if(confirm('This will delete current employee and then automatic add employee.\nDo you want to continue?')) window.location.href = 'cgi-bin/autoAddOrgEmp.asp?orgNo=<%=request.queryString("orgNo")%>&t=<%=orgType%>'">
              <input type="button" value="Back" onClick="javascript: window.location.href='orgLst.asp'">
            </td>
          </tr>
          <tr> 
            <td colspan="2" nowrap height=""><img src="../imgs/spacer.gif" height="4"></td>
          </tr>
          <!--########## EMPLOYEE TABLE ##########-->
          <tr> 
            <td colspan="2" nowrap valign="top"><b><font color="#000000">Child 
              organization (one level only):</font></b> </td>
          </tr>
          <!--########## CHILD TABLE ##########-->
          <tr> 
            <td colspan="2" nowrap valign="top"> 
              <table awidth="100%" border="0" cellspacing="0" cellpadding="0" width="100%">
                <tr> 
                  <td nowrap valign="top" bgcolor="#0000FF"> 
                    <table awidth="100%" border="0" cellspacing="1" cellpadding="2" width="100%">
                      <tr class="hdrRow" align="center"> 
                        <td nowrap awidth="1%"><b><font color="#000000">No</font></b></td>
                        <td nowrap awidth="1%"><b><font color="#000000">Org No</font></b></td>
                        <td nowrap awidth="1%"><b><font color="#000000">Name</font></b></td>
                        <td><b><font color="#000000">Desc</font></b></td>
                        <!--<td nowrap awidth="1%"><b><font color="#000000"></font></b></td>-->
                        <td nowrap awidth="1%"><b><font color="#000000">&nbsp;</font></b></td>
                      </tr>
                      <%
									sqlOra = "select lv, vOrgNo, vName, vDesc from (select level as lv, vOrgNo as vOrgNo2 from orgTree start with vOrgNo = '" & request.queryString("orgNo") & "' connect by vParent = prior vOrgNo), vOrgMst where vOrgNo = vOrgNo2 and lv = 2 order by lv"
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
                        <td nowrap awidth="1%" align="center"><font color="#000000"><%=i%></font></td>
                        <td nowrap awidth="1%" align="center"><font color="#000000"><%=rsOra("vOrgNo")%></font></td>
                        <td nowrap awidth="1%"><font color="#000000"><%=rsOra("vName")%></font></td>
                        <td><font color="#000000"><%=rsOra("vDesc")%></font></td>
                        <!--<td align="center" nowrap awidth="1%"><input type="button" value="Edit..."></td>-->
                        <td align="center" nowrap awidth="1%"> 
                          <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this child from tree?\n\nOrg No: <%=rsOra("vOrgNo") & "\nName: " & rsOra("vName") & "\nDesc: " & replace(replace(rsOra("vDesc"), "'", "\'"), QUOTE, "\" & QUOTE)%>')) { window.location.href = 'cgi-bin/delChild.asp?orgNo=<%=request.queryString("orgNo")%>&childNo=<%=rsOra("vOrgNo")%>' }">
                        </td>
                      </tr>
                      <%
												i = i + 1
												rsOra.moveNext
											loop
										else
									%>
                      <tr> 
                        <td nowrap align="center" colspan="5" bgcolor="#FFFFFF"><br>
                          <font class="alertMsg"><b>Child organization not found!</b></font><br>
                          &nbsp;</td>
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
          <!--########## CHILD TABLE ##########-->
        </table>
      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" awidth="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" awidth="17" height="18"></td>
      <td background="../imgs/border/book/bg.gif">&nbsp;</td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif"></td>
      <td align="center" background="../imgs/border/book/bg.gif"> 
        <input type="button" value="Add child organization..." onClick="javascript: window.location.href = 'addChildFrm.asp?orgNo=<%=request.queryString("orgNo")%>&orgName=<%=server.urlEncode(request.queryString("orgName") & "")%>&orgDesc=<%=server.urlEncode(request.queryString("orgDesc") & "")%>'">
      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
    </tr>
    <tr> 
      <td align="right" valign="top"><img src="../imgs/border/book/blc.gif" awidth="17" height="18"></td>
      <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" awidth="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/brc.gif" awidth="17" height="18"></td>
    </tr>
  </table>
</form>
<!--#include file="../inc/copyRight.inc"-->
<%hidePageLoading%>
</body>
</html>
