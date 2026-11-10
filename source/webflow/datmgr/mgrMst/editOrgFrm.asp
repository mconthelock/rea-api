<!--#include virtual="/sys/inc/system.inc"-->
<!--#include virtual="/sys/inc/as400.inc"-->
<!--#include virtual="/form/inc/system.inc"-->
<!--#include virtual="/form/inc/adovbs.inc"-->
<!--#include virtual="/form/inc/db.inc"-->
<!--#include virtual="/form/inc/connUtil.inc"-->
<!--#include virtual="/form/inc/authen.inc"-->
<!--#include virtual="/form/inc/misc.inc"-->
<!--#include virtual="/sys/inc/pageUtil.inc"-->
<!--#include file="../inc/datMgr.inc"-->
<!--include file="inc/wta.inc"-->
<!--include file="inc/db.inc"-->

<%
Dim rsOra, sqlOra
Dim i, pc, acode(1000)
Dim hName, hCode
set rsOra = createObject("adodb.recordset")

sqlOra = "select sposition, pos.*, a.sname, a.sempno, a.sseccode " _
& " from " _
& " amec.pposition, amec.aemployee a" _
& " , (select distinct level, possubTree.* from possubTree " _
& " start with " _
& " cStart = '1' connect by prior vposcode = vparent) pos " _
& " where pposition.sposcode = pos.vposcode " _
& " and a.sseccode='" & request("orgNo") & "'" _
& " and a.sposcode=pposition.sposcode " _
&" and a.cstatus='1'" _
& " order by level "

'#############	кр Section Manager	################
sqlOra="select sempno, sname from amec.aemployee, orgpos" _ 
& " where vorgno='" & request("orgNo") & "' and vposno='30'" _
& " and vempno=sempno"
'response.write sqlOra
'response.end
rsOra.open sqlOra, connOra
if not rsOra.EoF then
	hName = rsOra("sname")
	hCode = rsOra("sempno")
end if
rsOra.close
'#############	кр Section Manager	################

sqlOra = "select sq.*, a.sempno, a.sname, a.sposcode, a.sseccode, pos.sposition " _
& " from amec.aemployee a, amec.pposition pos," _
& " (select sq.*, level from sequenceOrg sq" _
& " start with vorgno1='" & request("orgNo") & "'" _
& " and sposcode1='30'" _
& " connect by prior empno=headno) sq" _
& " where a.sposcode=pos.sposcode " _
& " and sq.empno=a.sempno" _
& " and sq.sposcode=a.sposcode" _
& " and sq.vorgno=a.sseccode" _
& " and a.cstatus='1'"
'order by a.sname

sqlOra = "select orgpos.*,org.orgno, org.orgname, a.sname from orgpos, amec.aemployee a " _
& " , (select sseccode orgno, ssec orgname from amec.psection union" _
& " select sdepcode orgno, sdept orgname from amec.pdepartment union" _
& " select sdivcode orgno, sdiv orgname from amec.pdivision) org" _
& " where org.orgno = orgpos.vorgno "  _
& " and org.orgno='" & request("orgNo") & "'" _
& " and a.sempno = orgpos.vempno"
'response.write sqlOra & "<br>"
'response.end
rsOra.open sqlOra, connOra
%>
<html>
<head>
<!--#include virtual="/form/inc/metatag.inc"-->
<link href="/form/css/common.css" rel="stylesheet" type="text/css">
<title></title>
</head>
<body leftMargin="0" topMargin="0" marginHeight="0" marginWidth="0">

<form name="frmMain" method="get" action="">
  <table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
    <tr> 
      <td align="right" valign="bottom"><img src="../../imgs/border/book/tlc.gif" width="17" height="18"></td>
      <td background="../../imgs/border/book/t.gif" width="100%" align="center"><img src="../../imgs/border/book/t.gif" width="17" height="18"></td>
      <td align="left" valign="bottom"><img src="../../imgs/border/book/trc.gif" width="17" height="18"></td>
    </tr>
    <tr> 
      <td background="../../imgs/border/book/l.gif" align="right">&nbsp;</td>   
      <td bgcolor="#AACCEE" align="center"><b><font color="#000000">Subordinate tree master</font></b></td>
      <td background="../../imgs/border/book/r.gif" align="left">&nbsp;</td>
    </tr>
    <tr> 
      <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif" width="17" height="18"></td>
      <td bgcolor="#0000FF"> 
        <table width="100%" border="0" cellspacing="1" cellpadding="2">
          <tr class="hdrRow"> 
            <!--<td nowrap width="1%" align="center"><b>No</b></td>-->
            <!--<td nowrap align="center" width="0">
              <p><b><font color="#000000">Code</font></b></p>
              </td>-->
            <td nowrap align="center" width="100%" colspan=3><b><font color="#000000"> EmpCode-Ename&nbsp;(<%=rsOra("vorgno")%>)</font></b></td>
          </tr>
          <%
					pc = rsOra.EOF

					if not rsOra.EOF then
						i = 1
						do while not rsOra.EOF
			%>
          <tr bgcolor="#FFFFFF">		  
            <td nowrap align="left" width="1%"><img src="../../imgs../../imgs/user.gif"><font color=#000000>&nbsp;(<%=rsOra("vempno")%>)&nbsp;</font></td>
            <td nowrap align="left" width="0"><font color="#0000FF"><%= rsOra("sname")%></font></td>
			<td nowrap align="center" width="0"><input type="button" value="Edit..." onClick="javascript: window.location.href = 'eposf.asp?posNo=<%=rsOra("vposno")%>&orgno=<%=rsOra("orgNo")%>&headno=<%=rsOra("vempno")%>'"></td>
          </tr>			
				<%
							rsOra.moveNext
						loop
					else
			%>
          <tr> 
            <td align="center" colspan="4" bgcolor="#FFFFFF"><br>
              <font class="alertMsg"><b>Data not found!</b></font><br>
              &nbsp;</td>
          </tr>
          <%
					end if
					
					rsOra.close
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
            <td background="../../imgs/border/book/l.gif" align="right"><img src="../../imgs/border/book/l.gif"></td>
            <td align="center" background="../../imgs/border/book/bg.gif"> 
        <input type="button" value="Back" onClick="javascript: window.location.href = '/form/datmgr/index.asp'">
            </td>
            <td background="../../imgs/border/book/r.gif" align="left"><img src="../../imgs/border/book/r.gif"></td>
        </tr>
        <tr> 
            <td align="right" valign="top"><img src="../../imgs/border/book/blc.gif" width="17" height="18"></td>
            <td background="../../imgs/border/book/b.gif"><img src="../../imgs/border/book/b.gif" width="17" height="18"></td>
            <td align="left" valign="bottom"><img src="../../imgs/border/book/brc.gif" width="17" height="18"></td>
        </tr>
    </table>
</form>
<!--#include virtual="/sys/inc/copyRight.inc"-->

</body>
</html>


<%
set rsOra = nothing
%>