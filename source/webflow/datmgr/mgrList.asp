<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/sys/inc/connUtil.inc"-->
<!--#include virtual="/sys/inc/userUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->
<!--#include virtual="/sys/inc/report.inc"-->
<!--#include virtual="/sys/inc/reportData.inc"-->
<!--#include virtual="/sys/inc/misc.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->

<%
Dim rsOra, sqlOra
Dim pageSize, pageCount, absolutePage, recordCount, currentRecord  'Page/Report control variable
Dim whereClause, i
Dim fromClause
set rsOra = createObject("adodb.recordset")

'Set page size
if request.queryString("ps") = ""  then
	pageSize = 20
else
	pageSize = cint(request.queryString("ps"))
end if

if pageSize < 0 then
	pageSize = 20
end if

'Set absolute page
if request.queryString("ab") = "" then
	absolutePage = 1
else
	absolutePage = cint(request.queryString("ab"))
end if

if absolutePage < 0 then
	absolutePage = 1
end if

whereClause = "a.sEmpNo = sa.empNo and b.sEmpNo = sa.headNo " _
& "and sa.vOrgNo1 in ('" & request.queryString("selSec") & "', '" & request.queryString("selDep") & "', '" & request.queryString("selDiv") & "') " _
& "and ap.sPosCode = sa.sPosCode " _
& "and bp.sPosCode = sb.sPosCode1 " _
& "and sb.empNo = sa.empNo and sb.sPosCode = sa.sPosCode and sb.vOrgNo = sa.vOrgNo and sb.cCo = sa.cCo " _
& "and sb.headNo = sa.headNo and sb.sPosCode1 = sa.sPosCode1 and sb.vOrgNo1 = sa.vOrgNo1 and sb.cCo1 = sa.cCo1 " 

fromClause = "(select sseccode orgno, ssec orgname from amec.psection union " _ 
& "select sdepcode orgno, sdept orgname from amec.pdepartment union "_
& "select sdivcode orgno, sdiv orgname from amec.pdivision)"
%>

<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title></title>
<script language="JavaScript" src="/form/js/window.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftMargin="0" topMargin="0" marginWidth="0" marginHeight="0">
<%showPageLoading%>
<form name="frmMain" method="get">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr valign="bottom"> 
      <td align="right"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr> 
            <td bgcolor="#FFCC00" width="1%" nowrap align="center"><b><font color="#000000">No</font></b></td>
            <td bgcolor="#FFCC00" width="1%" nowrap align="center"><b><font color="#000000">Emp 
              No</font></b></td>
            <td bgcolor="#FFCC" align="center"> <font color="#000000"><b>Name</b></font></td>
            <td bgcolor="#FFCC" align="center" width="1%"><b><font color="#000000">Position</font></b></td>
            <td bgcolor="#FFCC" align="center" width="1%" nowrap><b><font color="#000000">Co-Pos</font></b></td>
            <td bgcolor="#FFCC00" nowrap width="1%" align="center"><b><font color="#000000">Mgr 
              No</font></b></td>
            <td bgcolor="#FFCC00" nowrap width="1%" align="center"><b><font color="#000000">Name</font></b></td>			  
            <td bgcolor="#FFCC00" nowrap width="1%" align="center"><b><font color="#000000">Org 
              No</font></b></td>
            <td bgcolor="#FFCC00" nowrap width="1%" align="center"><b><font color="#000000">Position</font></b></td>
            <td bgcolor="#FFCC00" width="1%" nowrap><b><font color="#000000">Co-Pos</font></b></td>
            <td bgcolor="#FFCC00" width="1%">&nbsp; </td>
            <td bgcolor="#FFCC00" width="1%">&nbsp;</td>
          </tr>
          <%
		  sqlOra = "select " _
		  & "sa.empNo, sa.vOrgNo, sa.cCo, sa.sposcode, sa.headno, sa.vOrgNo1, sa.sposcode1, sa.cCo1, a.sname, ap.sPosition as sPosition1, bp.sPosition as sPosition2, b.sname headname, orgname" _
		  & " from " _
		  & "sequenceOrg sa, sequenceOrg sb, amec.aEmployee a, amec.aEmployee b, amec.pPosition ap, amec.pPosition bp " _
		  & "," & fromClause &" where " & whereClause & " and sa.vOrgNo1 = orgno order by sa.headNo, sName"
		  'response.write sqlOra & "<br>"
		  'response.end
		  rsOra.open sqlOra, connOra

			'Goto start record
			recordCount = 1
			
			do while not rsOra.EOF and recordCount  < absolutePage * pageSize - pageSize + 1
				recordCount = recordCount + 1
				rsOra.moveNext
			loop
		
			currentRecord = absolutePage * pageSize - pageSize + 1
		
			i = currentRecord
			
		  if rsOra.EOF then
		  %>
          <tr> 
            <td bgcolor="#FFFFFF" align="center" colspan="11"><font class="alertMsg"><b><br>
              Data not found!<br>
              </b></font></td>
          </tr>
          <%
		  else
			do while not rsOra.EOF and i < absolutePage * pageSize + 1
		%>
          <tr bgcolor="#FFFFCC" valign="top"> 
            <td width="1%" nowrap align="center"><font color="#000000"><%=i%></font></td>
            <td width="1%" nowrap align="center"><font color="#000000"><a href="javascript: var winUsrInfo = window.open('/form/usrInfo.asp?uid=<%=rsOra("empNo")%>', 'winUsrInfo', WIN_CONF_REMARK); winUsrInfo.focus();"><%=rsOra("empNo")%></a></font></td>
            <td><font color="#000000"><%=rsOra("sName")%></font></td>
            <td width="1%"><font color="#000000"><%=rsOra("sPosition1")%></font></td>
            <td align="center" width="1%"><font color="#000000"> 
              <%if rsOra("cCo") <> "0" then response.write "Yes" else response.write "No" end if%>
              </font></td>
            <td nowrap width="1%" align="center"><font color="#000000"><a href="javascript: var winUsrInfo = window.open('/form/usrInfo.asp?uid=<%=rsOra("headNo")%>', 'winUsrInfo', WIN_CONF_REMARK); winUsrInfo.focus();"><%=rsOra("headNo")%></a></font></td>
           <td nowrap width="1%"><font color="#000000"><%=rsOra("headname")%></font></td>						  			
            <td nowrap width="1%" align="center"><font color="#000000"><!--<a href="javascript: var winUsrInfo = window.open('/form/usrInfo.asp?uid=<%'=rsOra("headNo")%>', 'winUsrInfo', WIN_CONF_REMARK); winUsrInfo.focus();"><%'=rsOra("vOrgNo1")%>--><%=rsOra("orgname")%><!--</a>--></font></td>
            <td nowrap width="1%"><font color="#000000"><%=rsOra("sPosition2")%></font></td>
            <td align="center" width="1%"><font color="#000000"> 
              <%if rsOra("cCo1") <> "0" then response.write "Yes" else response.write "No" end if%>
              </font></td>
            <td align="center" width="1%"> 
              <input type="button" value="Edit..." onClick="javascript: window.location.href = 'editMgrFrm.asp?e=<%=rsOra("empNo") & "&o=" & rsOra("vOrgNo") & "&p=" & rsOra("sPosCode") & "&c=" & rsOra("cCo") & "&h=" & rsOra("headNo") & "&o1=" & rsOra("vOrgNo1") & "&p1=" & rsOra("sPosCode1") & "&c1=" & rsOra("cCo1") & "&selDiv=" & request("selDiv") & "&selDep=" & request("selDep") & "&selSec=" & request("selSec")%>'">
            </td>
            <td align="center" width="1%"> 
              <input type="button" value="Delete" onClick="javascript: if(confirm('Do you want to delete this record?')) {window.location.href = 'cgi-bin/delMgr.asp?e=<%=rsOra("empNo") & "&o=" & rsOra("vOrgNo") & "&p=" & rsOra("sPosCode") & "&c=" & rsOra("cCo") & "&h=" & rsOra("headNo") & "&o1=" & rsOra("vOrgNo1") & "&p1=" & rsOra("sPosCode1") & "&c1=" & rsOra("cCo1") & "&selDiv=" & request("selDiv") & "&selDep=" & request("selDep") & "&selSec=" & request("selSec")%>'}">
            </td>
          </tr>
          <%
			  	i = i + 1
		  		rsOra.moveNext
		  	loop
		  end if
		  
		  rsOra.close
		  %>
        </table>
      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/bg.gif" align="center"> 
        <%
		'##### Display page navigator #####
		sqlOra = "select count(*) as recordCount " _
		& "from sequenceOrg sa, sequenceOrg sb, amec.aEmployee a, amec.aEmployee b, amec.pPosition ap, amec.pPosition bp " _
		& "where " & whereClause
		'response.write sqlOra & "<br>"
		rsOra.open sqlOra, connOra
		
		'		recordCount = currentRecord - 1 
		'recordCount = rsOra("recordCount")
		'Place Here
		displayPageNavigator absolutePage, rsOra("recordCount"), pageSize, request.serverVariables("SCRIPT_NAME") & "?no=" & request.queryString("no") & "&orgNo=" & request.queryString("orgNo") & "&y=" & request.queryString("y") & "&hDisplayData=1" & "&txtUid=" & request.queryString("txtUid") & "&txtUName=" & request.queryString("txtUName") & "&selSec=" & request.queryString("selSec") & "&selDep=" & request.queryString("selDep") & "&selDiv=" & request.queryString("selDiv") & "&search=" & request.queryString("search") & "&orgName=" & server.urlEncode(request.queryString("orgName") & "") & "&orgDesc=" & server.urlEncode(request.queryString("orgDesc") & "") & "&ab="
		rsOra.close
		%>
      </td>
      <td align="left" background="../imgs/border/book/r.gif"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif">&nbsp;</td>
      <td background="../imgs/border/book/bg.gif" align="center">&nbsp;</td>
      <td align="left" background="../imgs/border/book/r.gif">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif">&nbsp;</td>
      <td background="../imgs/border/book/bg.gif" align="center"> 
        <input type="button" name="Submit" value="Back" onClick="javascript: window.location.href = 'orgFrm.asp'">
      </td>
      <td align="left" background="../imgs/border/book/r.gif">&nbsp;</td>
    </tr>
    <!--
    <tr valign="top"> 
      <td align="right" background="../imgs/border/book/l.gif"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/bg.gif" align="center"> 
        <input type="button" value="Back" onClick="javascript: window.location.href = 'orgFrm.asp'">
      </td>
      <td align="left" background="../imgs/border/book/r.gif"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
	-->
    <tr valign="top"> 
      <td align="right"><img src="../imgs/border/book/blc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/b.gif"><img src="../imgs/border/book/b.gif" width="17" height="18"></td>
      <td align="left"><img src="../imgs/border/book/brc.gif" width="17" height="18"></td>
    </tr>
  </table>
</form>
<br>
<!--#include virtual="/sys/inc/copyRight.inc"-->
<br>
&nbsp; 
<%hidePageLoading%>
</body>
</html>

<%
set rsOra = nothing
%>