<%
Option Explicit
%>

<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/misc.inc"-->

<%
Dim rsOra, sqlOra
Dim i, orgType

set rsOra = createObject("adodb.recordset")
%>
<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
<title>Add organization</title></head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="frmMain" method="get" action="cgi-bin/addChild.asp">
	<input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
    
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr>
      <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Add child 
        to tree<br>
        "<%=request.queryString("orgNo")%>: <%=request.queryString("orgName")%> 
        (<%=request.queryString("orgDesc")%>)"</font></b></td>
      <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#FFFFFF"> 
        <table width="100%" border="0" cellspacing="2" cellpadding="1">
          <tr valign="top"> 
            <td width="1%" nowrap><b><font color="#000000">Organization:</font></b></td>
            <td></td>
          </tr>
          <!--########## ORGANIZATION LIST ##########-->
          <tr> 
            <td colspan="2" align="center"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td bgcolor="#0000FF" align="center"> 
                    <table width="100%" border="0" cellspacing="1" cellpadding="2">
                      <tr class="hdrRow" align="center"> 
                        <td nowrap width="1%"><b><font color="#000000">Type</font></b></td>
                        <td width="1%" nowrap><b><font color="#000000">&nbsp;</font></b></td>
                        <!--<td width="1%" nowrap><b><font color="#000000">No</font></b></td>-->
                        <td width="1%" nowrap><b><font color="#000000">Org No</font></b></td>
                        <td nowrap><b><font color="#000000">Name</font></b></td>
                        <td nowrap><b><font color="#000000">Description</font></b></td>
                      </tr>
                      <%
'										sqlOra = "select vOrgMst.*, vType from vOrgMst, orgType where vOrgNo not in ('00', '" & request.queryString("orgNo") & "') and vOrgNo not in (select distinct vOrgNo from orgTree where vParent = '" & request.queryString("orgNo") & "') and orgType.cType = vOrgMst.cType order by vOrgMst.cType desc, vOrgNo"

										sqlOra = "select vOrgMst.*, vType from vOrgMst, orgType where vOrgNo not in ('00', '" & request.queryString("orgNo") & "') and vOrgNo not in (select distinct vOrgNo from orgTree start with cStart = '1' connect by vParent = prior vOrgNo) and orgType.cType = vOrgMst.cType order by vOrgMst.cType desc, vOrgNo"
										'response.write sqlOra & "<br>"
										rsOra.open sqlOra, connOra
										
										if not rsOra.EOF then
											i = 1
											
											do while not rsOra.EOF
												'response.write "<input type=" & QUOTE & "checkbox" & QUOTE & " name=" & QUOTE & "chkChild" & QUOTE & " value=" & QUOTE & rsOra("vOrgNo") & QUOTE & ">" & rsOra("vOrgNo") & ": " & rsOra("vName") & " (" & rsOra("vDesc") & ")</option><br>"
										%>
                      <tr bgcolor="<%
												if i mod 2 = 0 then
													response.write "#FFCC00"
												else
													response.write "#FFFFCC"
												end if
												%>"> 
                        <td align="center" nowrap width="1%"><font color="#000000"> 
                          <%
												if orgType <> rsOra("vType") then
													orgType = rsOra("vType")
													response.write orgType
												else
													response.write "&nbsp;"	
												end if
												%>
                          </font></td>
                        <td align="center" width="1%"> 
                          <input type="checkbox" name="chkChild" value="<%=rsOra("vOrgNo")%>">
                        </td>
                        <!--<td align="center" width="1%"><font color="#000000"><%=i%></font></td>-->
                        <td align="center" width="1%"><font color="#000000"><%=rsOra("vOrgNo")%></font></td>
                        <td nowrap width="1%"><font color="#000000"><%=rsOra("vName")%></font></td>
                        <td><font color="#000000"><%=rsOra("vDesc")%></font></td>
                      </tr>
                      <%
												i = i + 1
												rsOra.moveNext

												if not rsOra.EOF then
													if orgType <> rsOra("vType") then
														response.write "<tr><td bgcolor=" & QUOTE & "#FFFFFF" & QUOTE & " colspan=" & QUOTE & "5" & QUOTE & ">&nbsp;</td></tr>"
													end if
												end if												
											loop
										else
										%>
                      <tr> 
                        <td align="center" nowrap width="1%" colspan="5"><font color="#000000">Organization 
                          not found!</font> </td>
                      </tr>
                      <%
										end if
										
										rsOra.close
										%>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <!--########## ORGANIZATION LIST ##########-->
        </table>
      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
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
        <input type="button" value="Back" onClick="javascript: window.location.href='editOrgFrm.asp?orgNo=<%=request.queryString("orgNo")%>'">
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
