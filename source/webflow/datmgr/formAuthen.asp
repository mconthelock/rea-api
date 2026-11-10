<!--#include virtual="/sys/inc/system.inc"-->
<!--#include file="../inc/system.inc"-->
<!--#include file="../inc/adovbs.inc"-->
<!--#include file="../inc/db.inc"-->
<!--#include file="../inc/connUtil.inc"-->
<!--#include virtual="/form/inc/formUtil.inc"-->

<%



Dim requesterNo, requester, inputerNo, inputer
Dim mode, flowStatus
Dim rsOra, rsOra2, sqlOra
Dim no, orgNo, y, y2, runNo, apv, remark, requestTo
Dim pageSize, pageCount, absolutePage, recordCount, currentRecord  'Page/Report control variable
Dim authNo, authName, authDesc
Dim i
Dim whereClause, orderClause
Dim haveData
Dim tmpWhere, tmpFrom

haveData = false

set rsOra = createObject("adodb.recordset")
set rsOra2 = createObject("adodb.recordset")

'Get mode
if session("uid") = request.queryString("empNo") then
	apv = request.queryString("empNo")
else
	apv = session("uid")
end if

mode = getMode(request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("y2"), request.queryString("runNo"), apv)

flowStatus = getFlowStatus(connOra, request.queryString("no"), request.queryString("orgNo"), request.queryString("y"), request.queryString("y2"), request.queryString("runNo"))

'Set page size
if request.queryString("ps") = ""  then
	pageSize = 10
else
	pageSize = cint(request.queryString("ps"))
end if

if pageSize < 0 then
	pageSize = 10
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

'########## GENERATE WHERE CLAUSE ##########
'whereClause = " sEmpNo in (select distinct vEmpNo from orgPos) and cStatus <> '0'"
whereClause = "cStatus <> '0'"

if request.queryString("search") = "1" then
	if request.queryString("txtUid") <> "" then
		if inStr(request.queryString("txtUid"), "*") > 0 or inStr(request.queryString("txtUid"), "?") > 0 then
			whereClause = whereClause & " and sEmpNo like '" & replace(replace(request.queryString("txtUid"), "*", "%"), "?", "_") & "'"
		else
			whereClause = whereClause & " and sEmpNo like '%" & request.queryString("txtUid") & "%'"
		end if
	end if

	if request.queryString("txtUName") <> "" then
		if inStr(request.queryString("txtUName"), "*") > 0 or inStr(request.queryString("txtUName"), "?") > 0 then
			whereClause = whereClause & " and sName like '" & replace(replace(ucase(request.queryString("txtUName")), "*", "%"), "?", "_") & "'"
		else
			whereClause = whereClause & " and sName like '%" & ucase(request.queryString("txtUName")) & "%'"
		end if
	end if

	if request.queryString("selSec") <> "" then
		whereClause = whereClause & " and sSecCode = '" & request.queryString("selSec") & "'"
	end if
	
	if request.queryString("selDep") <> "" then
		whereClause = whereClause & " and sDepCode = '" & request.queryString("selDep") & "'"
	end if
	
	if request.queryString("selDiv") <> "" then
		whereClause = whereClause & " and sDivCode = '" & request.queryString("selDiv") & "'"
	end if
	
	if request.queryString("selPos") <> "" then
		whereClause = whereClause & " and sPosCode = '" & request.queryString("selPos") & "'"
	end if

	if request.queryString("selAuth") <> "" then
		tmpFrom = ", usrAuth" 
		tmpWhere = " and cAuthno = '" & request.queryString("selAuth") & "' and aemployee.sempno = usrAuth.vempno " _
		& " and nfrmno = '" & request.queryString("no") & "' and vorgno='" & request.queryString("orgNo") & "'" _
		& " and cyear = '" & request.queryString("y") & "'"
'		& " and usrAuth.vorgno = aemployee.sseccode(+) " _ 
'		& " and usrAuth.vorgno = aemployee.sdepcode(+) " _ 
'		& " and usrAuth.vorgno = aemployee.sdivcode(+) " 
	end if
	
else
	'whereClause = "sSecCode = '" & request.queryString("orgNo") & "' or sDepCode = '" & request.queryString("orgNo") & "' or sDivCode = '" & request.queryString("orgNo") & "'"
	whereClause = "1 = 1"
end if

whereClause = whereClause & " and cStatus <> '0'"
orderClause = "order by sEmpNo"
'########## END GENERATE WHERE CLAUSE ##########
%><html>
<head> 
<!--#include virtual="/form/inc/metatag.inc"-->

<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/form.js"></script>
<script language="JavaScript" src="../js/window.js"></script>
<!--<script language="JavaScript" src="../js/eMailRegFrm.js"></script>--></head>
<body bgcolor="#FFFFFF" text="#000000" leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0"> 
<!--#include file="../inc/flow.inc"-->
<!--#include file="../inc/flowUtil.inc"-->
<!--#include file="../inc/hisUtil.inc"-->
<!--#include file="../inc/report.inc"-->
<!--#include file="../inc/reportData.inc"-->
<!--#include file="../inc/misc.inc"-->
<!--#include file="../inc/authen.inc"-->
<form name="frmMain" method="post" action="cgi-bin/saveAuthen.asp">
  <input type="hidden" name="no" value="<%=request.queryString("no")%>">
  <input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
  <input type="hidden" name="y" value="<%=request.queryString("y")%>">
  <input type="hidden" name="txtUid" value="<%=request.queryString("txtUid")%>">
  <input type="hidden" name="txtUName" value="<%=request.queryString("txtUName")%>">
  <input type="hidden" name="selSec" value="<%=request.queryString("selSec")%>">
  <input type="hidden" name="selDep" value="<%=request.queryString("selDep")%>">
  <input type="hidden" name="selDiv" value="<%=request.queryString("selDiv")%>">
  <input type="hidden" name="selPos" value="<%=request.queryString("selPos")%>">  
  <input type="hidden" name="search" value="<%=request.queryString("search")%>">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/t.gif" width="100%" align="center"><img src="../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right">&nbsp;</td>
      <td align="center" bgcolor="#99CCFF"><b><font color="#000000">User's form 
        authorization</font></b></td>
      <td background="../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
  <%
	sqlOra = "select to_char(nNo) as nNo, vOrgNo, cYear, vName, vDesc, vFormPage, vAuthPage from formMst where nNo =  '" & request.queryString("no") & "' and vOrgNo = '" & request.queryString("orgNo") & "' and cYear = '" & request.queryString("y") & "'"
	rsOra.open sqlOra, connOra
	
	if not rsOra.EOF then
	%>
  <!--########## FORM DATA ##########-->
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#FFFFFF"> 
        <table awidth="100%" border="0" cellpading="0" cellspacing="0" cellpadding="0">
          <tr valign="top" bgcolor="#0000FF"> 
            <td colspan="2" nowrap> 
              <table awidth="100%" border="0" cellspacing="1" cellpadding="2" width="100%">
                <tr> 
                  <td bgcolor="#FFCC00" awidth="1%"><b><font color="#000000">Form 
                    No:</font></b></td>
                  <td bgcolor="#FFFFCC"><font color="#000000"><%=rsOra("nNo")%></font></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFCC00" awidth="1%"><b><font color="#000000">Organization:</font></b></td>
                  <td bgcolor="#FFFFCC"><font color="#000000"> 
                    <%
			sqlOra = "select * from vOrgMst where vOrgNo = '" & request.queryString("orgNo") & "' order by cType desc, vOrgNo"
			rsOra2.open sqlOra, connOra
			
			if not rsOra.EOF then
				response.write rsOra("vOrgNo") & ": " & rsOra2("vName") & " (" & rsOra2("vDesc") & ")"
			else
				response.write "<font color=" & QUOTE & "#000000" & QUOTE & ">Organization not found!</font>"
			end if
			
			rsOra2.close
			%>
                    </font></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFCC00" awidth="1%"><b><font color="#000000">Name:</font></b></td>
                  <td bgcolor="#FFFFCC"><font color="#000000"><%=rsOra("vName")%></font></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFCC00" awidth="1%"><b><font color="#000000">Descript:</font></b></td>
                  <td bgcolor="#FFFFCC"><font color="#000000"><%=rsOra("vDesc")%></font></td>
                </tr>
              </table>
            </td>
          </tr>
          <%rsOra.close%>
          <%else%>
          <tr> 
            <td width="100%" align="center" bgcolor="#FFFFFF" colspan="2"><br>
              <font class="alertMsg"><b>Form not found!</b></font><br>
              &nbsp;</td>
          </tr>
          <%
			rsOra.close
		end if
	%>
          <!--########## END FORM DATA ##########-->
          <!--########## AUTHORITIES ##########-->
          <%
		sqlOra = "select usrAuth.*, sEmpNo, sName " _
		& "from " _
		& "usrAuth, amec.aEmployee "_
		& "where " _
		& whereClause _
		& " and usrAuth.nFrmNo(+) = '" & request.queryString("no") & "' " _
		& "and usrAuth.vOrgNo(+) = '" & request.queryString("orgNo") & "' " _
		& "and usrAuth.cYear(+) = '" & request.queryString("y") & "' and vEmpNo(+) = sEmpNo " _
		& orderClause
		'response.write sqlOra & "<br>"
		
		sqlOra = "select sEmpNo, sName, sSecCode, sDepCode, sDivCode, cAuthNo " _
		& "from amec.aEmployee,usrAuth "  
		sqlOra = sqlOra & "where " & whereClause _
		& tmpWhere _
		& " and nFrmNo(+) = '" & request.queryString("no") & "' and vOrgNo(+) = '" & request.queryString("orgNo") & "' and cYear(+) = '" & request.queryString("y") & "' and vEmpNo(+) = sEmpNo " _
		& orderClause
		'response.write sqlOra
'					response.end
		rsOra.open sqlOra, connOra
	%>
          <tr> 
            <td colspan="2"><img src="../imgs/spacer.gif" height="4"></td>
          </tr>
          <tr> 
            <td colspan="2"><b><font color="#000000">Employee:</font></b></td>
          </tr>
          <tr> 
            <td colspan="2"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td bgcolor="#0000FF"> 
                    <table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">
                      <tr bgcolor="#99CCFF"> 
                        <td width="1%" align="center" nowrap rowspan="2"><b><font color="#000000">No</font></b></td>
                        <td width="1%" align="center" nowrap rowspan="2"><b><font color="#000000">Emp-No</font></b></td>
                        <td width="100%" align="center" nowrap rowspan="2"><b><font color="#000000">Emp-Name</font></b></td>
                        <td colspan="3" align="center" nowrap><b><font color="#000000">Authority</font></b></td>
                      </tr>
                      <tr> 
                        <td align="center" nowrap bgcolor="#99CCFF"><img src="/sys/imgs/person06t.gif" alt="Form administrator"></td>
                        <td align="center" nowrap bgcolor="#99CCFF"><img src="/sys/imgs/person06bt.gif" alt="Form key operator"></td>
                        <td align="center" nowrap bgcolor="#99CCFF"><img src="/sys/imgs/person06gt.gif" alt="Form user"></td>
                      </tr>
                      <%
			'Read auth master
			sqlOra = "select * from formAuth order by cNo"
			rsOra2.open sqlOra, connOra
			
			if not rsOra2.EOF then
				if not rsOra.EOF then
					haveData = true
					
					do while not rsOra2.EOF
						authNo = rsOra2("cNo")
						authName = rsOra2("vName")
						authDesc = rsOra2("vDesc")
						
						rsOra2.moveNext
					loop

					i = 1
					
					'Goto start record
					recordCount = 1
					
					do while not rsOra.EOF and recordCount  < absolutePage * pageSize - pageSize + 1
						recordCount = recordCount + 1
						rsOra.moveNext
					loop
				
					currentRecord = absolutePage * pageSize - pageSize + 1
				
					i = currentRecord
					
					do while not rsOra.EOF and i < absolutePage * pageSize + 1
			%>
                      <input type="hidden" name="empList" value="<%=rsOra("sEmpNo")%>">
                      <tr bgcolor="#<%
				if i mod 2 = 0 then
					response.write "FFCC00"
				else
					response.write "FFFFCC"
				end if
				%>"> 
                        <td width="1%" nowrap align="center"><font color="#000000"><%=i%></font></td>
                        <td width="1%" align="center" nowrap><font color="#000000"><%=rsOra("sEmpNo")%></font></td>
                        <td width="100%" nowrap><font color="#000000"><%=rsOra("sName")%></font></td>
                        <td width="1%" nowrap align="center"> 
                          <input type="radio" name="optAuth<%=rsOra("sEmpNo")%>" value="<%=rsOra("sEmpNo") & "," & 	AUTH_FORM_ADMIN%>" <%
				if rsOra("cAuthNo") = AUTH_FORM_ADMIN then
					response.write "checked"
				end if%>>
                        </td>
                        <td width="1%" align="center"> 
                          <input type="radio" name="optAuth<%=rsOra("sEmpNo")%>" value="<%=rsOra("sEmpNo") & "," & AUTH_KEY_OPR%>" <%
				if rsOra("cAuthNo") = AUTH_KEY_OPR then
					response.write "checked"
				end if%>>
                        </td>
                        <td width="1%" align="center"> 
                          <input type="radio" name="optAuth<%=rsOra("sEmpNo")%>" value="<%=rsOra("sEmpNo") & "," & AUTH_FORM_USER%>" <%
				if rsOra("cAuthNo") = AUTH_FORM_USER or isNull(rsOra("cAuthNo")) then
					response.write "checked"
				end if%>>
                        </td>
                      </tr>
                      <%
						i = i + 1
						rsOra.moveNext
					loop
				end if
				
				rsOra.close
			%>
                      <%
				else
			%>
                      <tr> 
                        <td bgcolor="#FFFFFF" align="center" colspan="6"><br>
                          <font class="alertMsg"><b>Authorities master not found!</b></font><br>
                          &nbsp;</td>
                      </tr>
                      <%
				end if
				
				rsOra2.close
			%>
                    </table>
                  </td>
                </tr>
                <input type="hidden" name="totalEmp" value="<%=i%>">
                <!--########## END AUTHORITIES ##########-->
              </table>
            </td>
          </tr>
        </table>
      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
    <!--########## PAGE NAVIGATOR ##########-->
    <%if haveData then%>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td abackground="../imgs/ctlbg.jpg" align="center" bgcolor="#FFFFFF"> 
        <%
	'##### Display page navigator #####
	sqlOra = "select count(*) as recordCount " _
	& "from amec.aEmployee " 
	if request.queryString("selAuth") <> "" then
		sqlOra = sqlOra & tmpFrom
	end if
	sqlOra = sqlOra &  " where " & whereClause _
	& tmpWhere  _
	& orderClause
	'response.write sqlOra & "<br>"
	'response.end
	rsOra.open sqlOra, connOra, adOpenStatic, adLockReadOnly, adCmdText
	
	'		recordCount = currentRecord - 1 
	'recordCount = rsOra("recordCount")
	'Place Here
	displayPageNavigator absolutePage, rsOra("recordCount"), pageSize, request.serverVariables("SCRIPT_NAME") & "?no=" & request.queryString("no") & "&orgNo=" & request.queryString("orgNo") & "&y=" & request.queryString("y") & "&hDisplayData=1" & "&txtUid=" & request.queryString("txtUid") & "&txtUName=" & request.queryString("txtUName") & "&selSec=" & request.queryString("selSec") & "&selDep=" & request.queryString("selDep") & "&selDiv=" & request.queryString("selDiv") & "&selPos=" & request.queryString("selPos") & "&search=" & request.queryString("search") & "&orgName=" & server.urlEncode(request.queryString("orgName") & "") & "&orgDesc=" & server.urlEncode(request.queryString("orgDesc") & "") & "&selAuth=" & server.urlEncode(request.queryString("selAuth")& "") & "&ab="
	rsOra.close
	%>
      </td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif"></td>
    </tr>
    <%end if%>
    <!--########## END PAGE NAVIGATOR ##########-->
    <tr> 
      <td background="../imgs/border/book/l.gif" align="left"><img src="../imgs/border/book/l.gif" width="17" height="18"></td>
      <td background="../imgs/border/book/bg.gif"><img src="../imgs/spacer.gif" height="4"></td>
      <td background="../imgs/border/book/r.gif" align="left"><img src="../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../imgs/border/book/l.gif" align="right"><img src="../imgs/border/book/l.gif"></td>
      <td align="center" background="../imgs/border/book/bg.gif"> 
        <%if haveData then%>
        <input type="submit" name="act" value="Save change">
        <input type="reset" value="Reset">
        <%end if%>
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
<!--######### SEARCH FORM ##########-->
<table align="center" width="80%">
  <form name="frmSearchUsr" action="<%=request.serverVariables("SCRIPT_NAME")%>" mothod="get">
    <input type="hidden" name="no" value="<%=request.queryString("no")%>">
    <input type="hidden" name="orgNo" value="<%=request.queryString("orgNo")%>">
    <input type="hidden" name="y" value="<%=request.queryString("y")%>">
    <input type="hidden" name="search" value="1">
    <tr> 
      <td colspan="3"  abackground="../imgs/ctlbg.jpg"> 
        <hr size="1">
      </td>
    </tr>
    <tr abackground="../imgs/ctlbg.jpg"> 
      <td colspan="3"> 
        <table border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td colspan="2"><font class="FontText"><b><font color="#000000">Employee 
              search</font></b></font></td>
          </tr>
          <tr> 
            <td width="1%" nowrap><b><font class="FontText" color="#000000">Employee 
              ID:</font></b></td>
            <td> 
              <input type="text" name="txtUid" class="inputNormal" value="<%=request.queryString("txtUid")%>">
            </td>
          </tr>
          <tr> 
            <td width="1%" nowrap><b><font class="FontText" color="#000000">Name</font></b></td>
            <td> 
              <input type="text" name="txtUName" class="inputNormal" value="<%=request.queryString("txtUName")%>">
            </td>
          </tr>
          <tr> 
            <td width="1%" nowrap><b><font class="FontText" color="#000000">Section:</font></b></td>
            <td> 
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
          <tr> 
            <td width="1%" nowrap><b><font class="FontText" color="#000000">Department:</font></b></td>
            <td> 
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
          <tr> 
            <td width="1%" nowrap><b><font class="FontText" color="#000000">Division:</font></b></td>
            <td> 
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
          <tr> 
            <td width="1%" nowrap><b><font class="FontText" color="#000000">Position:</font></b></td>
            <td> 
              <%
		sqlOra = "select sPosCode, sPosition from amec.pPosition order by sPosition"
		rsOra.open sqlOra, connOra
		
		if not rsOra.EOF then
		%>
              <select name="selPos" class="inputNormal">
                <option value=""></option>
                <%do while not rsOra.EOF%>
                <option value="<%=rsOra("sPosCode")%>" <%if request.queryString("selPos") = rsOra("sPosCode") then response.write "selected" end if%>><%=rsOra("sPosition")%></option>
                <%
				rsOra.moveNext
			loop
			%>
              </select>
              <%
			else
				response.write "Position data not found!"
			end if
			
			rsOra.close
			%>
            </td>
          </tr>	  
          <tr> 
            <td width="1%" nowrap><b><font class="FontText" color="#000000">Authority:</font></b></td>
            <td> 
              <%
'		sqlOra = "select * from formAuth order by cNo"
'		rsOra.open sqlOra, connOra
'		if not rsOra.EOF then
		%>
              <select name="selAuth" class="inputNormal">
                <option value=""></option>
                <%'do while not rsOra.EOF%>
                <option value="<%=AUTH_KEY_OPR%>" <%if request.queryString("selAuth") = AUTH_KEY_OPR then response.write "selected" end if%>><% response.write "Key operator"%></option>
                <option value="<%=AUTH_FORM_ADMIN %>" <%if request.queryString("selAuth") = AUTH_FORM_ADMIN  then response.write "selected" end if%>><% response.write "Form administrator"%></option>
                <option value="<%=AUTH_FORM_USER %>" <%if request.queryString("selAuth") = AUTH_FORM_USER  then response.write "selected" end if%>><% response.write "Form user"%></option>
				
                <%
'				rsOra.moveNext
'			loop
			%>
              </select>
              <%
'			else
'				response.write "Division data not found!"
'			end if
		
'			rsOra.close
			%>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td colspan="3"  abackground="../imgs/ctlbg.jpg" align="right"><img src="/sys/imgs/spacer.gif" height="4"></td>
    </tr>
    <tr> 
      <td colspan="3"  abackground="../imgs/ctlbg.jpg" align="right"> 
        <input class="inputNormal" type="button" value="List all" onClick="javascript: window.location.href = '<%=request.serverVariables("SCRIPT_NAME")%>?no=<%=request.queryString("no") & "&orgNo=" & request.queryString("orgNo") & "&y=" & request.queryString("y")%>'">
        <input class="inputNormal" type="submit" value="Search">
      </td>
    </tr>
  </form>
</table>
<!--######### END SEARCH FORM ##########-->
<!--#include file="../inc/copyRight.inc"-->
</body>
</html>
