<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include file="../inc/misc.inc"-->
<!--#include file="../inc/userUtil.inc"-->
<!--#include file="../inc/formUtil.inc"-->

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
<table border="0" cellspacing="0" cellpadding="0" align="center" awidth="100%" width="100%">
  <tr> 
    <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" awidth="17" height="18"></td>
    <td background="../imgs/border/book/t.gif" awidth="100%" align="center" width="100%"><img src="../imgs/border/book/t.gif" awidth="17" height="18"></td>
    <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" awidth="17" height="18"></td>
  </tr>
  <tr> 
    <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" awidth="17" height="18"></td>
    <td bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td colspan="2" nowrap bgcolor="#AACCEE" align="center"><b><font color="#000000">Edit 
            report</font></b></td>
        </tr>
        <%
					sqlOra = "select to_char(nRepNo) as nRepNo, vOrgNo, cYear, vName, vDesc, vUrl, cStatus " _
					& "from " _
					& "rptMst " _
					& "where " _
					& "nRepNo =  '" & request("no") & "' and vOrgNo = '" & request("orgNo") & "' and cYear = '" & request("y") & "'"
					'response.write sqlOra
					rsOra.open sqlOra, connOra
					
					if not rsOra.EOF then
					%>
        <!--########## FORM DATA ##########-->
        <form name="frmMain" method="get" action="cgi-bin/editRpt.asp" enctype="multipart/form-data">
          <input type="hidden" name="no" value="<%=rsOra("nRepNo")%>">
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
									if rsOra2("vOrgNo") = request("orgNo") then
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
                        <td width="1%" nowrap bgcolor="#FFCC00"><b><font color="#000000">Organize:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <%showOrgList "name==selOrg&&selected==" & request("orgNo")%>
                        </td>
                      </tr>
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Name:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <input type="text" name="txtName" value="<%=rsOra("vName")%>" maxlength="60" size="35" title="Form's name">
                        </td>
                      </tr>
                      <tr valign="top"> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Description:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <textarea name="txtDesc" cols="40" title="Form's  description" rows="5"><%=rsOra("vDesc")%></textarea>
                        </td>
                      </tr>
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">URL:</font></b></td>
                        <td bgcolor="#FFFFCC"> 
                          <input type="text" name="txtFormPage" value="<%=rsOra("vUrl")%>" maxlength="1024" size="45">
                        </td>
                      </tr>
                      <!--
                      <tr>
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><font color="#000000"><b>Icon:</b></font></td>
                        <td bgcolor="#FFFFCC">
                          <input type="file" name="file">
                          <font color="#0000FF">(left blank to use default icon)</font> 
                        </td>
                      </tr>
					  -->
                      <tr> 
                        <td nowrap bgcolor="#FFCC00" awidth="1%" width="1%"><b><font color="#000000">Visibility</font></b></td>
                        <td bgcolor="#FFFFCC">
                          <input type="radio" name="cVisible" value="<%=RPT_STS_SHOW%>" <%if rsOra("cStatus") = RPT_STS_SHOW then response.write "checked" end if%>>
                          <font color="#000000">Show</font>&nbsp; 
                          <input type="radio" name="cVisible" value="<%=RPT_STS_HIDE%>" <%if rsOra("cStatus") = RPT_STS_HIDE then response.write "checked" end if%>>
                          <font color="#000000">Hide</font></td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2">&nbsp;</td>
          </tr>
          <%rsOra.close%>
          <tr align="center"> 
            <td colspan="2"> 
              <input type="submit" value="Save change" title="Save change to database">
              <input type="button" value="Authorization..." onClick="javascript: window.location.href = 'repAuthen.asp?no=<%=request("no")%>&orgNo=<%=request("orgNo")%>&y=<%=request("y")%>'">
              <!input type="button" value="Configuration...">
              <input type="button" value="Back" onClick="javascript: window.location.href='repLst.asp'">
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
            <input type="button" value="Back" onClick="javascript: window.location.href = 'frmLst.asp'">
          </td>
        </tr>
        <%
					rsOra.close
				end if
				%>
        <!--########## END FORM DATA ##########-->
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

