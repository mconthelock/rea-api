<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include virtual="/sys/inc/connUtil.inc"-->
<!--#include file="../inc/authen.inc"-->
<!--#include virtual="/sys/inc/misc.inc"-->
<!--#include virtual="/sys/inc/report.inc"-->
<!--#include virtual="/sys/inc/reportData.inc"-->
<!--#include virtual="/sys/inc/search.inc"-->
<!--#include virtual="/sys/inc/oraUtil.inc"-->
<!--#include virtual="/sys/inc/color.inc"-->

<%
Dim rsOra, sqlOra
Dim i
Dim pageSize, pageCount, absolutePage, recordCount, currentRecord  'Page/Report control variable
Dim whereClause, orderClause
Dim mode
Dim no, orgNo, y, rptName, empNo

set rsOra = createObject("adodb.recordset")

empNo = session("uid")

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

'########## GET SEL FORM ##########
if request.queryString("selScope") = "" then
	no = ""
	orgNo = ""
	y = ""
	rptName = ""
else
	no = left(request.queryString("selScope"), inStr(request.queryString("selScope"), "_")-1)
	orgNo = mid(request.queryString("selScope"), len(no)+2, 5)
	y = mid(request.queryString("selScope"), len(no + orgNo)+3, 2)
	rptName = mid(request.queryString("selScope"), len(no + orgNo + y) + 4, len(request.queryString("selScope")))
end if
'########## END GET SEL FORM ##########

'########## GENERATE WHERE CLAUSE ##########
whereClause = "where vEmpNo(+) = sEmpNo and cStatus <> '0' " _
& "and amec.pSection.sSecCode = amec.aEmployee.sSecCode " _
& "and amec.pDepartment.sDepCode = amec.aEmployee.sDepCode " _
& "and amec.pDivision.sDivCode = amec.aEmployee.sDivCode " _
& "and vEmpNo = amec.aEmployee.sEmpNo " _
& "and nFrmNo = '" & no & "' and vOrgNo = '" & orgNo & "' and cYear = '" & y & "' "

if no <> "" then
	whereClause = whereClause & "and nFrmNo(+) = '" & no & "' and vOrgNo(+) = '" & orgNo & "' and cYear(+) = '" & y & "' "
end if

if request.queryString("txtEmpNo") <> "" then
	whereClause = whereClause & " and sEmpNo like '" & getSearchStr(request.queryString("txtEmpNo")) & "'"
end if

if request.queryString("txtName") <> "" then
	whereClause = whereClause & " and lower(sName) like '" & getSearchStr(lCase(request.queryString("txtName"))) & "'"
end if

if request.queryString("selSec") <> "" then
	whereClause = whereClause & " and amec.aEmployee.sSecCode = '" & request.queryString("selSec") & "'"
end if

if request.queryString("selDep") <> "" then
	whereClause = whereClause & " and amec.aEmployee.sDepCode = '" & request.queryString("selDep") & "'"
end if

if request.queryString("selDiv") <> "" then
	whereClause = whereClause & " and amec.aEmployee.sDivCode = '" & request.queryString("selDiv") & "'"
end if
'########## END GENERATE WHERE CLAUSE ##########

'########## GENERATE ORDER CLAUSE ##########
orderClause = "order by amec.aEmployee.sEmpNo "
'########## END GENERATE ORDER CLAUSE ##########

'########## SET MODE ##########
if getAuthV2(session("sysauth"), AUTH_SYS_ADMIN) then
	mode = MODE_EDIT
else
	mode = MODE_VIEW
end if
'########## END SET MODE ##########
%>
<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"--> 
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title></title>
<script language="JavaScript" src="/sys/js/window.js"></script>
<script language="JavaScript">
function swapImg(imgObj, src1, src2)
{
	if(imgObj.src == src1)
	{
		imgObj.src = src2;
	}
	else
	{
		imgObj.src = src1;
	}
	imgObj.src = '../imgs/person6.gif';
}
</script>
</head>
<body text="#000000" leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0"> 
<center><b><font color="#000000">Report Authentication</font></b></center>
<form name="frmMain" method="post" action="cgi-bin/saveRptAuth.asp">
  <input type="hidden" name="txtEmpNo" value="<%=server.urlEncode(request.queryString("txtEmpNo"))%>">
  <input type="hidden" name="txtName" value="<%=server.urlEncode(request.queryString("txtName"))%>">
  <input type="hidden" name="selSec" value="<%=request.queryString("selSec")%>">
  <input type="hidden" name="selDep" value="<%=request.queryString("selDep")%>">
  <input type="hidden" name="selDiv" value="<%=request.queryString("selDiv")%>">
  <input type="hidden" name="scope" value="<%=request.queryString("selScope")%>">
  <input type="hidden" name="no" value="<%=no%>">
  <input type="hidden" name="orgNo" value="<%=orgNo%>">
  <input type="hidden" name="y" value="<%=y%>">
  
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/line/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/t.gif" width="100%" valign="bottom"><img src="../imgs/border/line/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/line/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td align="right" valign="bottom" background="../imgs/border/line/l.gif"><img src="../imgs/border/line/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/bg.gif" width="100%" valign="bottom" align="center"><b><font color="#000000">Employee 
        List<%if request.queryString("selScope") <> "" then response.write " of report <font color=#FF0000>" & QUOTE & "</font><font color=#0000FF>" & rptName & "</font><font color=#FF0000>" & QUOTE & "</font>" end if%></font></b></td>
      <td align="left" valign="bottom" background="../imgs/border/line/r.gif"><img src="../imgs/border/line/r.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td align="right" background="../imgs/border/line/l.gif"><img src="../imgs/border/line/l.gif" width="17" height="18"></td>
      <td bgcolor="#0000FF"> 
        <table width="100%%" border="0" cellspacing="1" cellpadding="2">
          <tr align="center" bgcolor="#CCCCCC"> 
            <td width="1%" nowrap rowspan="2"><b><font color="#000000">No</font></b></td>
            <td width="1%" nowrap rowspan="2"><b><font color="#000000">Emp No</font></b></td>
            <td width="100%" nowrap rowspan="2"><b><font color="#000000">Name</font></b></td>
            <td width="100%" nowrap rowspan="2"><b><font color="#000000">Sec No</font></b></td>
            <td width="100%" nowrap rowspan="2"><b><font color="#000000">Dep No</font></b></td>
            <td width="100%" nowrap rowspan="2"><b><font color="#000000">Div No</font></b></td>
            <td colspan="2" nowrap><b><font color="#000000">Authority</font></b></td>
          </tr>
          <tr align="center" bgcolor="#CCCCCC"> 
            <td width="1%" nowrap><img src="http://<%=request.serverVariables("SERVER_NAME")%>/form/imgs/reportItemread.gif" alt="Read"></td>
            <td width="1%" nowrap><img src="http://<%=request.serverVariables("SERVER_NAME")%>/form/imgs/reportItemnoread.gif" alt="Can't read"></td>
          </tr>
          <%
					sqlOra = "select " _
					& "distinct sEmpNo, sSec, sDept, sDiv, sName, cAuthNo " _
					& "from "
					
					sqlOra = sqlOra & "amec.aEmployee, amec.pSection, amec.pDepartment, amec.pDivision, rptUsrAuth "
					
					sqlOra = sqlOra & whereClause _
					& " " & orderClause
					'response.write sqlOra & "<br>"
					'response.end
					rsOra.open sqlOra, connOra
					
					i = 1
					
					do while not rsOra.EOF and i  < absolutePage * pageSize - pageSize + 1
						i = i + 1
						rsOra.moveNext
					loop
					
					if rsOra.EOF then
					%>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="13" align="center">&nbsp;<br>
              <font class="alertMsg">Data not found!</font><br>
              &nbsp;</td>
          </tr>
          <%
					else
						do while not rsOra.EOF and i < absolutePage * pageSize + 1
					%>
          <input type="hidden" name="empList" value="<%=rsOra("sEmpNo")%>">
          <tr bgcolor="#<%
		  if i mod 2 = 0 then
		  	response.write ROWA_COLOR
		  else
		  	response.write ROWB_COLOR
		  end if
		  %>"> 
            <td align="center"><font color="#000000"><%=i%></font></td>
            <td align="center"><font color="#000000"><a href="javascript: var winUsrInfo =  window.open('/<%=SYSTEM_FOLDER%>/usrInfo.asp?uid=<%=rsOra("sEmpNo")%>', 'winUsrInfo', 'width=600,height=400,location=no,ienubar=no,toolbar=no,scrollbars=no'); winUsrInfo.focus(); void(0)" onMouseOver="javascript: window.status = 'Click to see information of <%=rsOra("sEmpNo")%>'; return true" onMouseOut="javascript: window.status = '<%=SYSTEM_TITLE%>'; return true"><%=rsOra("sEmpNo")%></a></font></td>
            <td width="100%"><font color="#000000"><%=rsOra("sName")%></font></td>
            <td width="100%" align="center" nowrap><font color="#000000"><%=rsOra("sSec")%></font></td>
            <td width="100%" align="center" nowrap><font color="#000000"><%=rsOra("sDept")%></font></td>
            <td width="100%" align="center" nowrap><font color="#000000"><%=rsOra("sDiv")%></font></td>
            <td align="center" width="1%"> 
              <input type="radio" name="optAuth<%=rsOra("sEmpNo")%>" value="<%=AUTH_RPT_READ%>" <%if rsOra("cAuthNo") = AUTH_RPT_READ then response.write "checked" end if%>>
            </td>
            <td align="center" width="1%"> 
              <input type="radio" name="optAuth<%=rsOra("sEmpNo")%>" value="<%=AUTH_RPT_NOREAD%>" <%if rsOra("cAuthNo") = AUTH_RPT_NOREAD or isNull(rsOra("cAuthNo")) then response.write "checked" end if%>>
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
      <td align="left" background="../imgs/border/line/r.gif"><img src="../imgs/border/line/r.gif" width="17" height="18"></td>
    </tr>
    <!--########## PAGE NAVIGATOR ##########-->
    <tr> 
      <td background="../imgs/border/line/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/bg.gif" align="center"> 
        <%
			sqlOra = "select " _
			& "distinct count(*) as recordCount " _
			& "from "
			
			sqlOra = sqlOra & "amec.aEmployee, amec.pSection, amec.pDepartment, amec.pDivision, rptUsrAuth "
			
			sqlOra = sqlOra & whereClause _
			'response.write sqlOra & "<br>"
			rsOra.open sqlOra, connOra
			
			'		recordCount = currentRecord - 1 
			'recordCount = rsOra("recordCount")
			'Place Here
			displayPageNavigator absolutePage, rsOra("recordCount"), pageSize, request.serverVariables("SCRIPT_NAME") & "?hDisplayData=1" & "&selScope=" & server.urlEncode(request("selScope")) & "&txtEmpNo=" & request.queryString("txtEmpNo") & "&txtName=" & request.queryString("txtName") & "&selSec=" & request.queryString("selSec") & "&selDep=" & request.queryString("selDep") & "&selDiv=" & request.queryString("selDiv") & "&search=" & request.queryString("search") & "&ab="
			rsOra.close
			%>
      </td>
      <td align="left" background="../imgs/border/line/r.gif" valign="top">&nbsp;</td>
    </tr>
    <!--########## END PAGE NAVIGATOR ##########-->
    <%if mode = MODE_EDIT then%>
    <tr> 
      <td align="right" background="../imgs/border/line/l.gif" valign="top">&nbsp;</td>
      <td background="../imgs/border/line/bg.gif" align="center">&nbsp;</td>
      <td align="left" background="../imgs/border/line/r.gif" valign="top">&nbsp;</td>
    </tr>
    <tr> 
      <td align="right" background="../imgs/border/line/l.gif" valign="top"><img src="../imgs/border/line/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/bg.gif" align="center"> 
        <input type="submit" value="Save Change">
        <input type="reset" value="Reset">
        <input type="button" value="Back" onClick="javascript: window.location.href = 'index.asp'">
      </td>
      <td align="left" background="../imgs/border/line/r.gif" valign="top"><img src="../imgs/border/line/r.gif" width="17" height="18"></td>
    </tr>
    <%end if%>
    <tr> 
      <td align="right" valign="top"><img src="../imgs/border/line/blc.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/b.gif" valign="top"><img src="../imgs/border/line/b.gif" width="17" height="18"></td>
      <td align="left" valign="top"><img src="../imgs/border/line/brc.gif" width="17" height="18"></td>
    </tr>
  </table>
</form>
<!--########## MATERIAL SEARCH ##########-->
<form name="frmSearch" action="rptAuthen.asp" method="get">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="80%">
    <tr valign="bottom"> 
      <td align="right"><img src="../imgs/border/line/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/t.gif" width="100%"><img src="../imgs/border/line/t.gif" width="17" height="18"></td>
      <td align="left"><img src="../imgs/border/line/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td align="right" background="../imgs/border/line/l.gif"><img src="../imgs/border/line/l.gif" width="17" height="18"></td>
      <td bgcolor="#0000FF"> 
        <table width="100%%" border="0" cellspacing="1" cellpadding="2">
          <tr bgcolor="#CCCCCC"> 
            <td colspan="2"><b><font color="#000000">Employee Search</font></b></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td width="1%" nowrap bgcolor="#<%=HDR_CLR%>"><b><font color="#000000">Report:</font></b></td>
            <td bgcolor="#<%=DAT_CLR%>"><%
			sqlOra = "select to_number(nRepNo) as nRepNo, vOrgNo, cYear, vName from rptMst where cStatus = '" & RPT_STS_SHOW & "' order by vName"
			'response.write sqlOra & "<br>"
			rsOra.open sqlOra, connOra
			
			if rsOra.EOF then
				response.write "Report master not found!"
			else
			%>
			<select name="selScope">
				<option value="">&nbsp;</option>
			<%do while not rsOra.EOF%>
			<option value="<%=rsOra("nRepNo") & "_" & rsOra("vOrgNo") & "_" & rsOra("cYear") & "_" & rsOra("vName")%>" <%if rsOra("nRepNo") & "_" & rsOra("vOrgNo") & "_" & rsOra("cYear") = no & "_" & orgNo & "_" & y then response.write "selected" end if%>><%=rsOra("vName")%></option>
			<%
				rsOra.moveNext
			loop
			%>
			</select>
			<%
			end if
			
			rsOra.close
			%></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="1%" nowrap bgcolor="#<%=HDR_CLR%>"><b><font color="#000000">Employee 
              ID:</font></b></td>
            <td bgcolor="#<%=DAT_CLR%>"> 
              <input type="text" name="txtEmpNo" size="20" maxlength="12" value="<%=request.queryString("txtEmpNo")%>">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="1%" nowrap bgcolor="#<%=HDR_CLR%>"><b><font color="#000000">Name:</font></b></td>
            <td bgcolor="#<%=DAT_CLR%>"> 
              <input type="text" name="txtName" size="40" value="<%=request.queryString("txtName")%>">
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="1%" nowrap bgcolor="#<%=HDR_CLR%>"><b><font color="#000000">Section:</font></b></td>
            <td bgcolor="#<%=DAT_CLR%>"> 
              <%
				sqlOra = "select sSecCode, sSection from amec.pSection order by sSection"
				rsOra.open sqlOra, connOra
				
				if not rsOra.EOF then
				%>
              <select name="selSec" class="inputNormal">
                <option value=""></option>
                <%do while not rsOra.EOF%>
                <option value="<%=rsOra("sSecCode")%>" <%if request.queryString("selSec") = rsOra("sSecCode") then response.write "selected" end if%>><%=rsOra("sSection")%></option>
                <%
						rsOra.moveNext
					loop
					%>
              </select>
              <%
				else
				response.write "Section data not found!"
				end if
				
				rsOra.close
				%>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="1%" nowrap bgcolor="#<%=HDR_CLR%>"><b><font color="#000000">Department:</font></b></td>
            <td bgcolor="#<%=DAT_CLR%>"> 
              <%
				sqlOra = "select sDepCode, sDepartment from amec.pDepartment order by sDepartment"
				rsOra.open sqlOra, connOra
				
				if not rsOra.EOF then
				%>
              <select name="selDep" class="inputNormal">
                <option value=""></option>
                <%do while not rsOra.EOF%>
                <option value="<%=rsOra("sDepCode")%>" <%if request.queryString("selDep") = rsOra("sDepCode") then response.write "selected" end if%>><%=rsOra("sDepartment")%></option>
                <%
					rsOra.moveNext
					loop
					%>
              </select>
              <%
				else
				response.write "Department data not found!"
				end if
				
				rsOra.close
				%>
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="1%" nowrap bgcolor="#<%=HDR_CLR%>"><b><font color="#000000">Division:</font></b></td>
            <td bgcolor="#<%=DAT_CLR%>"> 
              <%
				sqlOra = "select sDivCode, sDivision from amec.pDivision order by sDivision"
				rsOra.open sqlOra, connOra
				
				if not rsOra.EOF then
				%>
              <select name="selDiv" class="inputNormal">
                <option value=""></option>
                <%do while not rsOra.EOF%>
                <option value="<%=rsOra("sDivCode")%>" <%if request.queryString("selDiv") = rsOra("sDivCode") then response.write "selected" end if%>><%=rsOra("sDivision")%></option>
                <%
						rsOra.moveNext
					loop
					%>
              </select>
              <%
				else
					response.write "Division data not found!"
				end if
				
				rsOra.close
				%>
            </td>
          </tr>
        </table>
      </td>
      <td align="left" background="../imgs/border/line/r.gif"><img src="../imgs/border/line/r.gif" width="17" height="18"></td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/line/l.gif"><img src="../imgs/border/line/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/bg.gif">&nbsp;</td>
      <td align="left" background="../imgs/border/line/r.gif"><img src="../imgs/border/line/r.gif" width="17" height="18"></td>
    </tr>
    <tr valign="top"> 
      <td align="right" background="../imgs/border/line/l.gif"><img src="../imgs/border/line/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/bg.gif" align="center"> 
        <input type="reset" value="Reset">
        <input type="button" value="List All" onClick="javascript: window.location.href = 'userList.asp'" name="button">
        <input type="submit" value="Search" name="search">
      </td>
      <td align="left" background="../imgs/border/line/r.gif"><img src="../imgs/border/line/r.gif" width="17" height="18"></td>
    </tr>
    <tr valign="top"> 
      <td align="right"><img src="../imgs/border/line/blc.gif" width="17" height="18"></td>
      <td background="../imgs/border/line/b.gif"><img src="../imgs/border/line/b.gif" width="17" height="18"></td>
      <td align="left"><img src="../imgs/border/line/brc.gif" width="17" height="18"></td>
    </tr>
  </table>
</form>
<!--########## END MATERIAL SEARCH ##########-->
<!--#include virtual="/sys/inc/copyRight.inc"-->
<br>
</body>
</html>
