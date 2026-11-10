<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/adovbs.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->
<%
Dim rsOra, sqlOra
Dim i, pc, pn, cp, empos
Dim whereclause
Dim hName, hCode
set rsOra = createObject("adodb.recordset")
'sqlOra = "select * from amec.pposition where sposcode = '" & request("p") & "' "
'sqlOra = "select empno,a.sempno,  a.sname, a.sposcode, pos.sposition from sequenceOrg, amec.aemployee a, amec.pposition pos" _
'& " where headno='" & request("headno") & "' and a.sempno=empno(+) and a.sseccode=vorgno(+) and pos.sposcode=a.sposcode"


'response.write sqlOra
'response.end
'rsOra.open sqlOra, connOra

'sqlOra="select sempno, sname from amec.aemployee a" _
'& " where sseccode='" & request("orgNo") & "' and a.sempno='" & request("headno") & "'"

sqlOra="select sempno , sname from amec.aemployee a, orgpos" _ 
& " where vorgno='" & request("orgNo") & "' and sempno='" & request("headno") & "'" _
& " and vempno=sempno"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra
if not rsOra.EoF then
	hName = rsOra("sname")
	hCode = rsOra("sempno")
end if
rsOra.close

select case request("posNo")
	case "02"
		whereclause=" and a.sposcode in ('10','40')"	
	case "10"
		whereclause=" and a.sposcode in ('11','20','21','30','35','50','55','43','41','40')"
	case "11"
		whereclause=" and a.sposcode in ('20','21','30','35','50','55','43','41','40')"
	case "20"
		whereclause=" and a.sposcode in ('21','30','35','50','55','43','41','40')"
	case "21"
		whereclause=" and a.sposcode in ('30','35','50','55','43','41','40')"		
	case "30"
		whereclause=" and a.sposcode in ('35','50','55','43','41','40')"
	case "50"
		whereclause=" and a.sposcode in ('55','41','40','43','62','60')"	
	case "55"
		whereclause=" and a.sposcode in ('41','43','60','62')"		
	case else
'		response.write request("posNo")
end select

'sqlOra="select empno,a.sempno,  a.sname, a.sposcode, pos.sposition" _
'& " , headno, a1.sname hName, pos1.sposition hPos" _
'& " from amec.aemployee a, amec.pposition pos, sequenceOrg sq" _
'& " , amec.aemployee a1, amec.pposition pos1" _
'& " where a.sposcode=pos.sposcode" _
'& " and sq.empno(+)=a.sempno" _
'& " and sq.sposcode(+)=a.sposcode" _
'& " and sq.vorgno(+)=a.sseccode" _
'& " and a.sseccode='" & request("orgNo") & "'" _
'& " and a.cstatus='1'" _
'& " and a1.sempno=headno" _
'& " and a1.sposcode=pos1.sposcode" _
'& whereclause _
'& " order by a.sname,sempno"	
'rsOra.open sqlOra, connOra

'if rsOra.EOF then
'	sqlOra="select a.sempno, a.sname, a.sposcode, pos.sposition" _
'	& " from amec.aemployee a, amec.pposition pos" _
'	& " where a.sposcode = pos.sposcode" _
'	& " and a.sseccode='" & request("orgNo") & "'" _
'	& whereclause _
'	& " order by a.sname,sempno"	
'end if
sqlOra = "select orgpos.*,org.orgno, org.orgname, a.sname, pos.sposition from orgpos, amec.aemployee a, amec.pposition pos " _
& " , (select sseccode orgno, ssec orgname from amec.psection union" _
& " select sdepcode orgno, sdept orgname from amec.pdepartment union" _
& " select sdivcode orgno, sdiv orgname from amec.pdivision) org" _
& " where org.orgno = orgpos.vorgno "  _
& " and a.sempno = orgpos.vempno" _
& " and orgpos.vposno >'" & request("posNo") & "'" _
& " and orgpos.vposno=pos.sposcode"  _
& " or org.orgno='" & request("orgNo") & "'" _
&  whereclause 

'if  request("posNo") <> "02" then
'	sqlOra = sqlOra & " or org.orgno='" & request("orgNo") & "'" 
'end if

response.write sqlOra
'response.end
'rsOra.close
rsOra.open sqlOra, connOra
%>
<html>
<head> 
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title>Edit organization</title></head>
<body text="#000000" leftMargin="0" topMargin="0" marginHeight="0" marginawidth="0"> 
<%=showPageLoading%> 
<form name="frmMain" method="get" action="ecpos.asp">
	<input type="hidden" name="headno" value="<%=request("headno")%>"> 
	<input type="hidden" name="orgno" value="<%=request("orgno")%>">
	<input type="hidden" name="posNo" value="<%=request("posNo")%>">
  <table border="0" cellspacing="0" cellpadding="0" align="center" awidth="100%" width="100%">
	<tr> 
	  <td align="right" valign="bottom"><img src="../../imgs/border/book/tlc.gif" width="17" height="18"></td>
	  <td background="../../imgs/border/book/t.gif" width="100%" align="center"><img src="../../imgs/border/book/t.gif" width="17" height="18"></td>
	  <td align="left" valign="bottom"><img src="../../imgs/border/book/trc.gif" width="17" height="18"></td>
	</tr>
	<tr> 
		<td background="../../imgs/border/book/l.gif" align="right">&nbsp;</td>   
		<td bgcolor="#AACCEE" align="center"><b><font color="#000000">Edit subordinate structure</font></b></td>
		<td background="../../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
    <tr> 
	  <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
	  <td bgcolor="#0000FF"> 
	  	<table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr> 
            <td width="16%" bgcolor="#FFCC00"><b><font color="#000000">Supervisor</font></b></td>
            <td width="84%" bgcolor="#FFFFCC"><font color="#00000">(<%=hCode%>)&nbsp;<%=hName%></font></td>
          </tr>
        </table>
	  </td>
      <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
   <tr> 
		<td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
		<td background="../../imgs/border/book/bg.gif">&nbsp;</td>
		<td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif"></td>
    </tr>
	<tr> 
		<td background="../../imgs/border/book/l.gif" align="right">&nbsp;</td>   
		<td bgcolor="#FFFFFF" align="left"><b><font color="#000000">Subordinate</font></b></td>
		<td background="../../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
	
    <tr> 
      <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr bgcolor="#FFCC00" align="center"> 
            <td bgcolor="#FFCC00" width="0%">&nbsp;</td>
            <td bgcolor="#FFCC00" width="0%"><b><font color="#000000">No</font></b></td>
            <td bgcolor="#FFCC00" nowrap><b><font color="#000000">Emp-Code</font></b></td>
            <td bgcolor="#FFCC00" nowrap><b><font color="#000000">Name</font></b></td>
            <td width="0%"><b><font color="#000000">Position</font></b></td>
            <td width="0%"><b><font color="#000000">Supervisor</font></b></td>
            <td width="0%"><b><font color="#000000">Supervisor-Position</font></b></td>
          </tr>
          <%
''		  sqlOra = "select " _
''		  & "sposcode, vposcode, sposition " _
''		  & "from " _
''		  & "amec.pposition, (select vposcode from evapostree start with vparent = '" & pc & "' connect by prior vposcode = vparent and level = '1') a " _
''		  & "where " _
''		  & "sposcode = vposcode(+) and sposcode <> '" & pc & "' " _
''		  & "order by " _
''		  & "sposcode "
''		  rsOra.open sqlOra, connOra
		  
			cp = rsOra.EOF
			
		  if rsOra.EOF then
		  %>
          <tr bgcolor="#FFFFCC"> 
            <td colspan="3" align="center"><font color="#FF0000"><b>Position master 
              not found!</b></font></td>
          </tr>
          <%
		  else
		  	i = 1
			
		  	do while not rsOra.EOF
		  %>
          <tr bgcolor="#FFFFCC"> 
            <td align="center"> 
              <input type="checkbox" name="c" value="<%=rsOra("vempno")%>">
            </td>
            <td align="center" width="0%"><font color="#000000"><%=i%></font></td>
            <td nowrap align=center width="0%"><font color="#000000"><%=rsOra("Vempno")%></font></td>
            <td nowrap><font color="#000000"><%=rsOra("sname")%></font></td>
            <td nowrap><font color="#000000"><%=rsOra("sposition")%></font></td>
            <td nowrap><font color="#000000">(<%'=rsOra("headno")%>)&nbsp;<%'=rsOra("hName")%></font></td>
            <td nowrap><font color="#000000"><%'=rsOra("hPos")%></font></td>
          </tr>
          <%
		  		rsOra.moveNext
				i = i + 1
		  	loop
		  end if
		  %>
        </table>
      </td>
      <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
      <tr> 
            <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
            <td background="../../imgs/border/book/bg.gif">&nbsp;</td>
            <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif"></td>
        </tr>
    <tr> 
     <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#FFFFFF" align="center"> 
        <%if not cp then%><input type="submit" value="Save"><%end if%>
        <input type="button" value="Cancel" onClick="javascript: window.location.href = 'posl.asp?orgNo=<%=request("orgno")%>'">
      </td>
     <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../../imgs/border/book/l.gif" width="2%"><img src="../../imgs/border/book/l.gif" awidth="17" height="18"></td>
      <td bgcolor="#FFFFFF"">&nbsp;</td>
      <td background="../../imgs/border/book/r.gif" align="left" width="2%"><img src="../../imgs/border/book/r.gif" awidth="17" height="18"></td>
    </tr>
<tr> 
      <td valign="top" width="2%" background="../../imgs/border/book/b.gif"><img src="../../imgs/border/book/blc.gif" awidth="17" height="18"></td>
      <td background="../../imgs/border/book/b.gif" ><img src="../../imgs/border/book/b.gif" awidth="17" height="18"></td>
      <td align="left" valign="bottom" width="2%"><img src="../../imgs/border/book/brc.gif" awidth="17" height="18"></td>
    </tr>
  </table>
  </form>
<!--#include virtual="/form/inc/copyRight.inc"-->
<%hidePageLoading%>
</body>
</html>

<%rsOra.close%>